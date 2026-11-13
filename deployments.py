import os
import re
import hashlib
import shutil
import tempfile
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
import plotly.express as px
import streamlit as st
from src.ui import download_row

SHEET = '10_DEPLOYMENTS'
COLUMNS = ['Deployment ID', 'Study ID', 'Study Name', 'Compound / Indication',
           'ARGES Hypothesis', 'FPI', 'LPI', 'DBL', 'IA', 'ARGES Analysis',
           'ICF', 'Protocol', 'Other', 'Endo', 'SRP', 'SRS', 'PML',
           'ARGES Personnel', 'Other Key Personnel', 'Next steps',
           'Source / Qualifications']
MISSING = {'', '-', '—', 'n/a', 'na', 'nan', 'none', 'not listed'}

def has_value(series):
    return ~series.fillna('').astype(str).str.strip().str.lower().isin(MISSING)

def prepare_deployments(data):
    df = data.get(SHEET, pd.DataFrame()).copy()
    df.columns = df.columns.astype(str).str.strip()
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    missing = set(COLUMNS) - set(df.columns)
    if missing:
        st.error('Missing deployment columns: ' + ', '.join(sorted(missing)))
        return pd.DataFrame(columns=COLUMNS)
    df = df.loc[has_value(df['Deployment ID']), COLUMNS].copy()
    for col in COLUMNS:
        df[col] = df[col].map(lambda x: '' if pd.isna(x) else
                              x.strftime('%Y-%m-%d') if hasattr(x, 'strftime') else str(x).strip())
    if df['Deployment ID'].duplicated().any():
        st.error('Deployment IDs must be unique.')
        return pd.DataFrame(columns=COLUMNS)
    return df.reset_index(drop=True)

def apply_study_filters(df, filtered, data):
    all_studies = data.get('01_STUDIES', pd.DataFrame())
    selected = filtered.get('01_STUDIES', pd.DataFrame())
    if 'Study ID' not in all_studies or 'Study ID' not in selected:
        return df
    all_ids = set(all_studies['Study ID'].dropna().astype(str).str.strip())
    ids = set(selected['Study ID'].dropna().astype(str).str.strip())
    if ids == all_ids:
        return df
    if not ids:
        return df.iloc[:0]
    return df[df['Study ID'].isin(ids)].copy()

def render(filtered, data, source=None, *, save_callback=None, source_label="Local workbook", allow_local_write=False):
    st.header('Deployments')
    st.caption('Clinical trial planning, timelines, documentation, key personnel and next steps for ARGES deployments.')
    full = prepare_deployments(data)
    if full.empty:
        st.info('No deployment information is available.')
        return
    df = apply_study_filters(full, filtered, data)
    st.subheader('Deployment Explorer')
    search = st.text_input('Search deployments', key='dep_search', placeholder='Study, compound, hypothesis, personnel...')
    with st.expander('Filter by column', expanded=True):
        columns = st.multiselect('Columns to filter', COLUMNS, key='dep_filter_columns')
        for col in columns:
            values = sorted(df[col].unique().tolist())
            options = ['(Blank)' if v == '' else v for v in values]
            chosen = st.multiselect(col, options, key='dep_filter_' + col)
            if chosen:
                df = df[df[col].isin(['' if v == '(Blank)' else v for v in chosen])]
    if search.strip():
        mask = df.apply(lambda s: s.str.contains(search.strip(), case=False, regex=False)).any(axis=1)
        df = df[mask]
    st.caption(f'{len(df)} of {len(full)} deployment records. Global study filters exclude records without a matching Study ID.')
    if df.empty:
        st.info('No deployments match the current filters.')
        return
    c1,c2,c3,c4 = st.columns(4)
    c1.metric('Deployment records', len(df))
    c2.metric('Protocols identified', int(has_value(df['Study ID']).sum()))
    c3.metric('DBL documented', int(has_value(df['DBL']).sum()))
    c4.metric('Next steps defined', int(has_value(df['Next steps']).sum()))
    st.subheader('Deployment Inventory')
    tabs = st.tabs(['Complete inventory', 'Timeline', 'Personnel', 'Documentation', 'Next steps'])
    with tabs[0]:
        st.dataframe(df, use_container_width=True, hide_index=True)
        download_row(df, 'filtered_deployments')
    groups = [(['Study Name','Study ID','FPI','LPI','DBL','IA','ARGES Analysis'], 'FPI: First Patient In | LPI: Last Patient In | DBL: Database Lock | IA: Interim Analysis'), (['Study Name','SRP','SRS','PML','ARGES Personnel','Other Key Personnel'], None), (['Study Name','ICF','Protocol','Other','Endo'], None), (['Study Name','ARGES Hypothesis','Next steps','Source / Qualifications'], None)]
    for tab,(cols,note) in zip(tabs[1:],groups):
        with tab:
            if note: st.caption(note)
            st.dataframe(df[cols], use_container_width=True, hide_index=True)
    render_editor(df, source, save_callback=save_callback, source_label=source_label, allow_local_write=allow_local_write)
    st.subheader('Deployment Portfolio')
    counts = df.apply(lambda row: classify(row, data), axis=1).value_counts().rename_axis('Indication').reset_index(name='Records')
    fig = px.bar(counts, x='Indication', y='Records', color='Indication', text='Records', color_discrete_map={'Ulcerative Colitis':'#E2231A', "Crohn's Disease":'#F46B62', 'Other / Unspecified':'#808899', 'Mixed / IBD':'#B39DDB'})
    fig.update_layout(template='plotly_dark', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', showlegend=False, height=300, margin=dict(l=10,r=10,t=15,b=10))
    st.plotly_chart(fig, use_container_width=True)
    st.caption('Disease from 01_STUDIES where matched; otherwise explicit indication or study-name evidence. Counts are deployment records.')


def disease_label(value):
    v = str(value).lower().replace('’', "'")
    uc = bool(re.search(r"\bulcerative colitis\b|\buc\b", v))
    cd = bool(re.search(r"\bcrohn(?:'s|s)? disease\b|\bcd\b", v))
    if uc and cd: return 'Mixed / IBD'
    if uc: return 'Ulcerative Colitis'
    if cd: return "Crohn's Disease"
    return 'Other / Unspecified'


def classify(row, data):
    studies = data.get('01_STUDIES', pd.DataFrame())
    sid = row['Study ID'].strip()
    if sid and {'Study ID', 'Disease'}.issubset(studies.columns):
        ids = studies['Study ID'].fillna('').astype(str).str.strip()
        matches = studies.loc[ids.eq(sid), 'Disease']
        if matches.empty and sid == 'CNTO1959CRD3001':
            matches = studies.loc[ids.str.startswith(sid + ' ('), 'Disease']
        labels = {disease_label(v) for v in matches} - {'Other / Unspecified'}
        if len(labels) == 1: return labels.pop()
        if len(labels) > 1: return 'Mixed / IBD'
    for value in (row['Compound / Indication'], row['Study Name']):
        label = disease_label(value)
        if label != 'Other / Unspecified': return label
    return 'Other / Unspecified'


def text_value(x):
    return '' if x is None or pd.isna(x) else x.strftime('%Y-%m-%d') if hasattr(x, 'strftime') else str(x).strip()


def save_changes(path, original, edited):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() != '.xlsx': raise ValueError('Only .xlsx sources are supported.')
    before = original.set_index('Deployment ID')
    after = edited.set_index('Deployment ID')
    if not before.index.is_unique or not after.index.is_unique or set(before.index) != set(after.index):
        raise ValueError('Deployment IDs must remain unique and unchanged.')
    changes = [(i,c,text_value(before.at[i,c]),text_value(after.at[i,c])) for i in before.index for c in before.columns if text_value(before.at[i,c]) != text_value(after.at[i,c])]
    if not changes: return 0
    if any(c in {'Study ID', 'Study Name'} for _,c,_,_ in changes):
        raise ValueError('Study identifiers are read-only.')
    for i,c,old,new in changes:
        if len(new)>32767: raise ValueError(f'{i}: {c} exceeds Excel cell limit.')
        if new.startswith('='): raise ValueError('Formula entry is not allowed.')
    lock = path.with_name(path.name + '.deployments.lock')
    fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    os.close(fd)
    tmp = None
    wb = None
    try:
        digest = hashlib.sha256(path.read_bytes()).digest()
        wb = load_workbook(path)
        ws = wb[SHEET]
        headers = {str(c.value).strip():c.column for c in ws[1] if c.value is not None}
        if not set(COLUMNS).issubset(headers): raise ValueError('Workbook headers changed.')
        rows = {}
        for r in range(2, ws.max_row+1):
            i = text_value(ws.cell(r,headers['Deployment ID']).value)
            if not i: continue
            if i in rows: raise ValueError('Duplicate Deployment ID in workbook.')
            rows[i] = r
        for i,c,old,new in changes:
            if i not in rows: raise ValueError(f'{i} no longer exists.')
            cell = ws.cell(rows[i],headers[c])
            if cell.data_type == 'f' or text_value(cell.value) != old:
                raise ValueError(f'Conflict in {i} / {c}. Reload before editing.')
            if new and c in ('FPI','LPI','DBL','IA','ARGES Analysis') and re.fullmatch(r'\d{4}-\d{2}-\d{2}',new):
                cell.value = datetime.strptime(new,'%Y-%m-%d')
                cell.number_format = 'yyyy-mm-dd'
            else:
                cell.value = new or None
                if new: cell.data_type = 's'
        with tempfile.NamedTemporaryFile(dir=path.parent,suffix='.xlsx',delete=False) as f: tmp=Path(f.name)
        wb.save(tmp)
        if hashlib.sha256(path.read_bytes()).digest()!=digest: raise ValueError('Workbook changed during save. Reload and retry.')
        backup = path.parent / 'arges_backups'
        backup.mkdir(exist_ok=True)
        shutil.copy2(path, backup / (path.stem+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.xlsx'))
        os.replace(tmp,path)
        return len(changes)
    finally:
        if wb: wb.close()
        if tmp and tmp.exists(): tmp.unlink()
        lock.unlink(missing_ok=True)


def render_editor(df, source, *, save_callback=None, source_label="Local workbook", allow_local_write=False):
    if st.session_state.pop('dep_saved', False): st.success('Deployment changes saved to the workbook.')
    with st.expander('Edit deployment records'):
        enabled = os.environ.get('ARGES_ENABLE_DEPLOYMENT_WRITES') == '1' and (save_callback is not None or (allow_local_write and source is not None))
        st.caption(f'Write target: {source_label}')
        st.caption('Edit visible records only. Deployment IDs are read-only. Save writes to the configured workbook; other records are preserved.')
        if not enabled: st.info('Persistent editing is disabled. An administrator must configure an authorized write target and enable workbook writes.')
        key = hashlib.sha256(df.to_json(orient='split').encode()).hexdigest()[:20]
        with st.form('dep_form_'+key):
            edited = st.data_editor(df, disabled=['Deployment ID', 'Study ID', 'Study Name'], num_rows='fixed', hide_index=True, use_container_width=True, key='dep_editor_'+key)
            submitted = st.form_submit_button('Save changes', disabled=not enabled)
        st.download_button('Download edited records (CSV)', edited.to_csv(index=False).encode('utf-8-sig'), file_name='deployments_edited.csv', mime='text/csv', key='dep_download_edits')
        if submitted:
            try:
                count = save_callback(df, edited) if save_callback is not None else save_changes(source, df, edited)
            except Exception as exc:
                st.error(f'Changes not saved: {exc}')
            else:
                if count:
                    st.cache_data.clear()
                    st.session_state['dep_saved'] = True
                    st.rerun()
                else: st.info('No changes to save.')
