import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inventory Hub", page_icon="📦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.stApp{background-color:#F1F8F4;}
section[data-testid="stSidebar"]{background-color:#1B5E20!important;}
section[data-testid="stSidebar"] *{color:#E8F5E9!important;}
.page-title{font-size:1.8rem;font-weight:700;color:#1B5E20;border-left:5px solid #4CAF50;padding-left:14px;margin-bottom:.3rem;}
.page-sub{font-size:.9rem;color:#4CAF50;padding-left:19px;margin-bottom:1.5rem;}
.success-box{background:#F0FDF4;border-left:4px solid #22C55E;padding:.7rem 1rem;border-radius:0 8px 8px 0;font-size:.88rem;color:#166534;margin-bottom:1rem;}
.warning-box{background:#FFF8E1;border-left:4px solid #F59E0B;padding:.7rem 1rem;border-radius:0 8px 8px 0;font-size:.88rem;color:#92400E;margin-bottom:1rem;}
</style>
""", unsafe_allow_html=True)

# ── Constants ──
ACTIVOS=['Carhartt','Champion','Duluth Trading','Harley-Davidson Motor Company, Inc.',
         'Kerusso inc.','STUSSY, Inc','Tiltworks','Under Armour','Uniforms','Vortex Optics']
INACTIVOS=['Adidas','CCM Hockey','Dallas Cowboys','Dickies','Fanatics Licensed Sports Group',
           'G&G Outfitters, INC','GFS','Grupo Beta','Jansport','Knigths Apparel','Outerstuff LTD',
           'Profile','Southern Tide','UA-GFS','Upward Sports','Golds Gym','Walls','SLD','Oakley']
NO_VMI=['Duluth Trading','Tiltworks','Vortex Optics']
CMAP_HN={'Regular':'Regulars','VMI':'VMI','Exceso':'Excess','Irregulares':'Irregulars','Obsoleto':'Obsolete'}
CMAP_TLP={}
CMAP_HN_CUSTOMER={}
TLP_ORDER=['Crazy Shirts','Dickies','Grupo Beta','INX PRINTS INC.','Outerstuff LTD','Renfro',
           'Simply Southern','Textiles La Paz','TLP OUTERSTUFF','TLP Vans','Under Armour','Vans','NOVUS','Timberland']

# ── Excel style helpers ──
NAVY='0D2B4E';NAVY_MID='1A4A7A';NAVY_LIGHT='2E6DA4';NAVY_PALE='D6E4F0'
NAVY_ALT='EBF3FB';WHITE='FFFFFF';GRAY_MID='D0D8E4';ACCENT='1B6CA8';SUBTOT='2E6DA4';GRAND='0D2B4E'

def brd():
    t=Side(style='thin',color=GRAY_MID)
    return Border(left=t,right=t,top=t,bottom=t)

def xh(ws,r,c,v,bg=NAVY,fg=WHITE,sz=10,b=True,ha='center'):
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=Font(bold=b,color=fg,name='Calibri',size=sz)
    cell.fill=PatternFill('solid',start_color=bg)
    cell.alignment=Alignment(horizontal=ha,vertical='center',wrap_text=True)
    cell.border=brd();return cell

def xd(ws,r,c,v,bg=None,b=False,ha='right'):
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=Font(bold=b,color='1A1A2E',name='Calibri',size=10)
    if bg:cell.fill=PatternFill('solid',start_color=bg)
    cell.alignment=Alignment(horizontal=ha,vertical='center')
    cell.border=brd()
    if isinstance(v,(int,float)) and v is not None:cell.number_format='#,##0'
    return cell

def xdf(ws,r,c,v,bg=None,hmode=False):
    fg=WHITE if hmode else ('1E6B3C' if v and v>0 else '8B1A1A' if v and v<0 else '1A1A2E')
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=Font(bold=True,color=fg,name='Calibri',size=10)
    if bg:cell.fill=PatternFill('solid',start_color=bg)
    cell.alignment=Alignment(horizontal='right',vertical='center')
    cell.border=brd()
    if isinstance(v,(int,float)) and v is not None:cell.number_format='+#,##0;-#,##0;0'
    return cell

# ── Program classification ──
def apply_program(df):
    df=df.copy();df['Program']=None
    df.loc[df['Box Usage'].astype(str).str.upper().str.contains('BLANK',na=False),'Program']='BLANKS'
    m2=df['Program'].isna()&df['Box Usage'].astype(str).str.upper().str.contains('PRINTED|EMBROIDERY',na=False)
    df.loc[m2,'Program']='PRINTED'
    df.loc[df['Program'].isna()&df['Option6'].isna(),'Program']='BLANKS'
    mo=df['Program'].isna()&df['Option6'].notna()
    df.loc[mo&df['Box\nStatus'].isin(['Packed','Picked']),'Program']='PRINTED'
    df.loc[mo&df['Box\nStatus'].isin(['Inventory','WIP']),'Program']='BLANKS'
    return df

# ── Honduras classification ──
def classify_honduras(carton_df, open_df):
    df=carton_df.copy()
    df['Quantity']=pd.to_numeric(df['Quantity'].astype(str).str.replace(',',''),errors='coerce').fillna(0)
    df['Customer Name']=df['Customer'].astype(str)
    df['Clasificacion']=None
    open_po=set(open_df['PONumber'].astype(str).str.strip().dropna())
    no_vmi=df['Customer Name'].isin(NO_VMI)
    df.loc[df['Is Second'].isin(['Second','Third']),'Clasificacion']='Irregulares'
    df.loc[(df['Order\nType']=='Locker Stock')&~no_vmi&df['Clasificacion'].isna(),'Clasificacion']='VMI'
    df.loc[(df['Customer']==12)&(df['Color\nDescription']=='PFD White')&~no_vmi&df['Clasificacion'].isna(),'Clasificacion']='VMI'
    df.loc[(df['Customer']==81)&df['Style'].astype(str).str.upper().str.endswith('-VMI')&~no_vmi&df['Clasificacion'].isna(),'Clasificacion']='VMI'
    mp=(df['Customer']==81)&df['Style'].astype(str).str.upper().str.contains('VMI-PULL')&~no_vmi
    df.loc[mp&~df['PONumber'].astype(str).str.strip().isin(open_po)&df['Clasificacion'].isna(),'Clasificacion']='VMI'
    df.loc[mp&df['PONumber'].astype(str).str.strip().isin(open_po)&df['Clasificacion'].isna(),'Clasificacion']='Regular'
    p1=df['PONumber'].astype(str).str.strip();p2=df['PONumbers'].astype(str).str.strip()
    df.loc[df['Clasificacion'].isna()&(p1.isin(open_po)|p2.isin(open_po)),'Clasificacion']='Regular'
    df.loc[df['Clasificacion'].isna()&df['Box Tag'].astype(str).str.contains('Excess',na=False),'Clasificacion']='Exceso'
    df.loc[df['Clasificacion'].isna()&(df['Box Tag']=='Obsolete'),'Clasificacion']='Obsoleto'
    df['Create Date']=pd.to_datetime(df['Create Date'],errors='coerce')
    cutoff=datetime.now().replace(year=datetime.now().year-1)
    ms=df['Order Status'].isin(['Complete','Void'])
    df.loc[df['Clasificacion'].isna()&ms&(df['Create Date']<=cutoff),'Clasificacion']='Obsoleto'
    df.loc[df['Clasificacion'].isna()&ms&(df['Create Date']>cutoff),'Clasificacion']='Exceso'
    df.loc[df['Clasificacion'].isna()&(df['Create Date']<=cutoff),'Clasificacion']='Obsoleto'
    df.loc[df['Clasificacion'].isna()&(df['Create Date']>cutoff),'Clasificacion']='Exceso'
    cut=df['Box Usage'].astype(str).str.upper().str.contains('CUT',na=False)
    cut_info=(cut.sum(),int(df.loc[cut,'Quantity'].sum()))
    df=df[~cut].copy()
    bs=df['Box\nStatus'];cl=df['Clasificacion']
    df['Type']=''
    df.loc[(cl=='Regular')&bs.isin(['Packed','Picked']),'Type']='Finished Goods'
    df.loc[(cl=='Regular')&bs.isin(['Inventory','WIP']),'Type']='Wip'
    df.loc[cl.isin(['VMI','Irregulares','Obsoleto','Exceso']),'Type']='Finished Goods'
    df=apply_program(df)
    df['Year']=df['Create Date'].dt.year.astype('Int64')
    return df,cut_info

# ── TLP classification ──
def classify_tlp(carton_df):
    df=carton_df.copy()
    df['Quantity']=pd.to_numeric(df['Quantity'].astype(str).str.replace(',',''),errors='coerce').fillna(0)
    df['Customer Name']=df['Customer'].astype(str)
    df['Clasificacion']=None
    df.loc[df['Is Second'].isin(['Second','Third']),'Clasificacion']='TLP Irregulars'
    df.loc[df['Clasificacion'].isna()&(df['Box Tag']=='Blanks Excess'),'Clasificacion']='TLP Blanks Excess'
    df.loc[df['Clasificacion'].isna()&(df['Box Tag']=='Printed Excess'),'Clasificacion']='TLP Printed Excess'
    df.loc[df['Clasificacion'].isna()&df['Box\nStatus'].isin(['Packed','Picked']),'Clasificacion']='TLP sin clasificacion'
    df.loc[df['Clasificacion'].isna(),'Clasificacion']='Wip'
    df=apply_program(df)
    df['Create Date']=pd.to_datetime(df['Create Date'],errors='coerce')
    df['Year']=df['Create Date'].dt.year.astype('Int64')
    return df

# ── Excel pivot writer ──
def write_pivot_sheet(ws,df,inv_cols,activos,inactivos):
    df=df.copy();df['Clas_Col']=df['Clasificacion'].map(CMAP_HN)
    p=df.pivot_table(index='Customer Name',columns='Clas_Col',values='Quantity',aggfunc='sum',fill_value=0)
    for c in inv_cols:
        if c not in p.columns:p[c]=0
    p=p[inv_cols];p['Grand Total']=p.sum(axis=1)
    cols=inv_cols+['Grand Total']
    ws.merge_cells('C1:H1');ws['C1']='Inventory Type'
    ws['C1'].font=Font(bold=True,name='Calibri',size=10,color=NAVY)
    ws['C1'].alignment=Alignment(horizontal='center')
    for ci,h in enumerate(['Clientes','Style CustomerName']+cols,1):xh(ws,2,ci,h)
    row=3
    for gn,gl in [('Activos',activos),('Inactivos',inactivos)]:
        clients=[c for c in gl if c in p.index]
        if not clients:continue
        first=True
        for i,cli in enumerate(clients):
            r=p.loc[cli];fill=NAVY_ALT if i%2==0 else None
            c1=ws.cell(row=row,column=1,value=gn if first else '')
            c1.font=Font(bold=True,color=WHITE,name='Calibri',size=10)
            c1.fill=PatternFill('solid',start_color=ACCENT)
            c1.alignment=Alignment(horizontal='center',vertical='center')
            c1.border=brd();first=False
            xd(ws,row,2,cli,bg=fill,ha='left')
            for ci,col in enumerate(cols,3):xd(ws,row,ci,int(r[col]) if r[col]!=0 else None,bg=fill)
            row+=1
        st_row=p.loc[clients].sum()
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
        xh(ws,row,1,f'{gn} Total',SUBTOT)
        for ci,col in enumerate(cols,3):xh(ws,row,ci,int(st_row[col]) if st_row[col]!=0 else None,SUBTOT)
        row+=1
    gt=p.sum()
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
    xh(ws,row,1,'Grand Total',GRAND)
    for ci,col in enumerate(cols,3):xh(ws,row,ci,int(gt[col]) if gt[col]!=0 else None,GRAND)
    ws.column_dimensions['A'].width=14;ws.column_dimensions['B'].width=38
    for ci in range(3,3+len(cols)):ws.column_dimensions[get_column_letter(ci)].width=15
    ws.freeze_panes='A3'

# ── Honduras Excel ──
def build_excel_hn(df,wk_prev=None,wk_label='WK13'):
    inv_cols=['Regulars','VMI','Excess','Irregulars','Obsolete']
    years=sorted(df['Year'].dropna().unique())
    wb=Workbook();wb.remove(wb.active)

    write_pivot_sheet(wb.create_sheet('Resumen Total'),df,inv_cols,ACTIVOS,INACTIVOS)
    write_pivot_sheet(wb.create_sheet('Finished Goods'),df[df['Type']=='Finished Goods'],inv_cols,ACTIVOS,INACTIVOS)
    write_pivot_sheet(wb.create_sheet('Wip'),df[df['Type']=='Wip'],inv_cols,ACTIVOS,INACTIVOS)

    # Antigüedad
    ws_age=wb.create_sheet('Antigüedad')
    for label,d in [('Finished Goods',df[df['Type']=='Finished Goods']),('Wip',df[df['Type']=='Wip'])]:
        p=d.pivot_table(index='Customer Name',columns='Year',values='Quantity',aggfunc='sum',fill_value=0)
        for yr in years:
            if yr not in p.columns:p[yr]=0
        p=p[[yr for yr in years]];p['Grand Total']=p.sum(axis=1)
        sr=ws_age.max_row+1 if ws_age.max_row>1 else 1
        ncols=2+len(years)+1
        ws_age.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=ncols)
        c=ws_age.cell(row=sr,column=1,value=f'  {label}')
        c.font=Font(bold=True,color=WHITE,name='Calibri',size=12)
        c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
        c.border=brd();ws_age.row_dimensions[sr].height=22;sr+=1
        ws_age.merge_cells(start_row=sr,start_column=3,end_row=sr,end_column=2+len(years))
        xh(ws_age,sr,1,'Clientes',NAVY_MID);xh(ws_age,sr,2,'Customer Name',NAVY_MID)
        xh(ws_age,sr,3,'Year of Creation',NAVY_MID)
        for ci in range(4,2+len(years)+1):
            c=ws_age.cell(row=sr,column=ci);c.fill=PatternFill('solid',start_color=NAVY_MID);c.border=brd()
        xh(ws_age,sr,2+len(years)+1,'Grand Total',NAVY_MID);sr+=1
        xh(ws_age,sr,1,'',NAVY_LIGHT);xh(ws_age,sr,2,'',NAVY_LIGHT)
        for yi,yr in enumerate(years):xh(ws_age,sr,3+yi,int(yr),NAVY_LIGHT)
        xh(ws_age,sr,3+len(years),'Grand Total',NAVY_LIGHT);sr+=1
        for gn,gl in [('Activos',ACTIVOS),('Inactivos',INACTIVOS)]:
            clients=[c for c in gl if c in p.index]
            if not clients:continue
            first=True
            for i,cli in enumerate(clients):
                r=p.loc[cli];rb=NAVY_ALT if i%2==0 else None
                c1=ws_age.cell(row=sr,column=1,value=gn if first else '')
                c1.font=Font(bold=True,color=WHITE,name='Calibri',size=10)
                c1.fill=PatternFill('solid',start_color=ACCENT)
                c1.alignment=Alignment(horizontal='center',vertical='center');c1.border=brd();first=False
                xd(ws_age,sr,2,cli,bg=rb,ha='left')
                for yi,yr in enumerate(years):xd(ws_age,sr,3+yi,int(r[yr]) if r[yr]!=0 else None,bg=rb)
                xd(ws_age,sr,3+len(years),int(r['Grand Total']),bg=NAVY_PALE,b=True);sr+=1
            st2=p.loc[clients].sum()
            ws_age.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=2)
            xh(ws_age,sr,1,f'{gn} Total',SUBTOT)
            for yi,yr in enumerate(years):xh(ws_age,sr,3+yi,int(st2[yr]) if st2[yr]!=0 else None,SUBTOT)
            xh(ws_age,sr,3+len(years),int(st2['Grand Total']),SUBTOT);sr+=1
        gt=p.sum()
        ws_age.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=2)
        xh(ws_age,sr,1,'Grand Total',GRAND)
        for yi,yr in enumerate(years):xh(ws_age,sr,3+yi,int(gt[yr]) if gt[yr]!=0 else None,GRAND)
        xh(ws_age,sr,3+len(years),int(gt['Grand Total']),GRAND)

    ws_age.column_dimensions['A'].width=14;ws_age.column_dimensions['B'].width=36
    for ci in range(3,3+len(years)+1):ws_age.column_dimensions[get_column_letter(ci)].width=10
    ws_age.column_dimensions[get_column_letter(3+len(years))].width=13;ws_age.freeze_panes='C4'

    # Damage Severity
    ws_dmg=wb.create_sheet('Damage Severity')
    d=df[df['Clasificacion']=='Irregulares'].copy()
    d['DS_Group']=d['Damage Severity'].apply(lambda x:'0, 1' if x in [0,1] else '2+')
    p=d.pivot_table(index='Customer Name',columns='DS_Group',values='Quantity',aggfunc='sum',fill_value=0)
    for c in ['0, 1','2+']:
        if c not in p.columns:p[c]=0
    p=p[['0, 1','2+']];p['Grand Total']=p.sum(axis=1)
    p=p.sort_values('Grand Total',ascending=False)
    ws_dmg.merge_cells('A1:E1')
    c=ws_dmg.cell(row=1,column=1,value='  Damage Severity — Irregulars Summary')
    c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
    c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
    c.border=brd();ws_dmg.row_dimensions[1].height=26
    ws_dmg.merge_cells('C2:D2')
    xh(ws_dmg,2,1,'Inventory Type',NAVY_MID);xh(ws_dmg,2,2,'Style CustomerName',NAVY_MID)
    xh(ws_dmg,2,3,'Damage Severity',NAVY_MID)
    ws_dmg.cell(row=2,column=4).fill=PatternFill('solid',start_color=NAVY_MID);ws_dmg.cell(row=2,column=4).border=brd()
    xh(ws_dmg,2,5,'Grand Total',NAVY_MID)
    xh(ws_dmg,3,1,'',NAVY_LIGHT);xh(ws_dmg,3,2,'',NAVY_LIGHT)
    xh(ws_dmg,3,3,'0, 1',NAVY_LIGHT);xh(ws_dmg,3,4,'2 +',NAVY_LIGHT);xh(ws_dmg,3,5,'Grand Total',NAVY_LIGHT)
    row=4;first=True
    for i,(cli,r) in enumerate(p.iterrows()):
        fill=NAVY_ALT if i%2==0 else None
        c1=ws_dmg.cell(row=row,column=1,value='Irregulars' if first else '')
        c1.font=Font(bold=True,color=WHITE,name='Calibri',size=10)
        c1.fill=PatternFill('solid',start_color=ACCENT)
        c1.alignment=Alignment(horizontal='center',vertical='center');c1.border=brd();first=False
        xd(ws_dmg,row,2,cli,bg=fill,ha='left')
        xd(ws_dmg,row,3,int(r['0, 1']) if r['0, 1']!=0 else None,bg=fill)
        xd(ws_dmg,row,4,int(r['2+']) if r['2+']!=0 else None,bg=fill)
        xd(ws_dmg,row,5,int(r['Grand Total']),bg=fill,b=True);row+=1
    gt=p.sum()
    ws_dmg.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
    xh(ws_dmg,row,1,'Grand Total',GRAND)
    xh(ws_dmg,row,3,int(gt['0, 1']) if gt['0, 1']!=0 else None,GRAND)
    xh(ws_dmg,row,4,int(gt['2+']) if gt['2+']!=0 else None,GRAND)
    xh(ws_dmg,row,5,int(gt['Grand Total']),GRAND)
    for col,w in zip('ABCDE',[16,28,14,14,14]):ws_dmg.column_dimensions[col].width=w
    ws_dmg.freeze_panes='C4'

    # Program
    ws_prog=wb.create_sheet('Program')
    dp=df[df['Type']=='Finished Goods'].copy()
    pp=dp.pivot_table(index='Clasificacion',columns='Program',values='Quantity',aggfunc='sum',fill_value=0)
    for c in ['BLANKS','PRINTED']:
        if c not in pp.columns:pp[c]=0
    pp=pp[['BLANKS','PRINTED']];pp['Grand Total']=pp.sum(axis=1)
    order=['Regular','VMI','Exceso','Irregulares','Obsoleto']
    pp=pp.reindex([c for c in order if c in pp.index]+[c for c in pp.index if c not in order])
    ws_prog.merge_cells('A1:E1')
    c=ws_prog.cell(row=1,column=1,value='  Program Summary — BLANKS vs PRINTED  (Finished Goods)')
    c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
    c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
    c.border=brd();ws_prog.row_dimensions[1].height=26
    ws_prog.merge_cells('B2:C2')
    xh(ws_prog,2,1,'Inventory Type',NAVY_MID);xh(ws_prog,2,2,'Program',NAVY_MID)
    ws_prog.cell(row=2,column=3).fill=PatternFill('solid',start_color=NAVY_MID);ws_prog.cell(row=2,column=3).border=brd()
    xh(ws_prog,2,4,'Grand Total',NAVY_MID)
    xh(ws_prog,3,1,'',NAVY_LIGHT);xh(ws_prog,3,2,'BLANKS',NAVY_LIGHT)
    xh(ws_prog,3,3,'PRINTED',NAVY_LIGHT);xh(ws_prog,3,4,'Grand Total',NAVY_LIGHT)
    for i,(clas,r) in enumerate(pp.iterrows()):
        fill=NAVY_ALT if i%2==0 else None
        xd(ws_prog,4+i,1,clas,bg=fill,ha='left')
        xd(ws_prog,4+i,2,int(r['BLANKS']) if r['BLANKS']!=0 else None,bg=fill)
        xd(ws_prog,4+i,3,int(r['PRINTED']) if r['PRINTED']!=0 else None,bg=fill)
        xd(ws_prog,4+i,4,int(r['Grand Total']),bg=fill,b=True)
    last=4+len(pp);gt=pp.sum()
    xh(ws_prog,last,1,'Grand Total',GRAND,ha='left')
    xh(ws_prog,last,2,int(gt['BLANKS']) if gt['BLANKS']!=0 else None,GRAND)
    xh(ws_prog,last,3,int(gt['PRINTED']) if gt['PRINTED']!=0 else None,GRAND)
    xh(ws_prog,last,4,int(gt['Grand Total']),GRAND)
    for col,w in zip('ABCD',[26,16,16,16]):ws_prog.column_dimensions[col].width=w
    ws_prog.freeze_panes='A4'

    # Comparativo
    if wk_prev:
        inv_types=['Regulars','VMI','Excess','Irregulars','Obsolete','Liability']
        df2=df.copy();df2['Clas_Col']=df2['Clasificacion'].map(CMAP_HN)
        wk13_p=df2[df2['Type']=='Finished Goods'].pivot_table(index='Customer Name',columns='Clas_Col',values='Quantity',aggfunc='sum',fill_value=0)
        for t in inv_types:
            if t not in wk13_p.columns:wk13_p[t]=0
        wk13_p['Liability']=0
        ws_comp=wb.create_sheet('Comparativo')
        ncols=2+len(inv_types)*3+3
        ws_comp.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
        c=ws_comp.cell(row=1,column=1,value=f'  Inventory Comparison — {wk_label} vs WK Anterior  (Finished Goods)')
        c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
        c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
        c.border=brd();ws_comp.row_dimensions[1].height=26
        xh(ws_comp,2,1,'Clientes',NAVY_MID);xh(ws_comp,2,2,'Customer Name',NAVY_MID)
        col=3
        for t in inv_types:
            ws_comp.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+2)
            xh(ws_comp,2,col,t,NAVY_MID);col+=3
        ws_comp.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+2)
        xh(ws_comp,2,col,'Grand Total',NAVY)
        xh(ws_comp,3,1,'',NAVY_LIGHT);xh(ws_comp,3,2,'',NAVY_LIGHT)
        col=3
        for t in inv_types:
            xh(ws_comp,3,col,'WK Ant.',NAVY_LIGHT);xh(ws_comp,3,col+1,wk_label,NAVY_LIGHT);xh(ws_comp,3,col+2,'Diff',NAVY_LIGHT);col+=3
        xh(ws_comp,3,col,'WK Ant.',NAVY);xh(ws_comp,3,col+1,wk_label,NAVY);xh(ws_comp,3,col+2,'Diff',NAVY)
        row=4;tot12={t:0 for t in inv_types};tot13={t:0 for t in inv_types}
        for gn,gl in [('Activos',ACTIVOS),('Inactivos',INACTIVOS)]:
            first=True
            for i,cli in enumerate(gl):
                w12=wk_prev.get(cli,{t:0 for t in inv_types})
                w13={t:int(wk13_p.loc[cli,t]) if cli in wk13_p.index and t in wk13_p.columns else 0 for t in inv_types}
                fill=NAVY_ALT if i%2==0 else None
                c1=ws_comp.cell(row=row,column=1,value=gn if first else '')
                c1.font=Font(bold=True,color=WHITE,name='Calibri',size=10)
                c1.fill=PatternFill('solid',start_color=ACCENT)
                c1.alignment=Alignment(horizontal='center',vertical='center');c1.border=brd();first=False
                xd(ws_comp,row,2,cli,bg=fill,ha='left')
                col=3;gt12=gt13=0
                for t in inv_types:
                    v12=w12[t];v13=w13[t];diff=v13-v12
                    xd(ws_comp,row,col,v12 if v12 else None,bg=fill)
                    xd(ws_comp,row,col+1,v13 if v13 else None,bg=fill)
                    xdf(ws_comp,row,col+2,diff if diff!=0 else None,bg=fill)
                    gt12+=v12;gt13+=v13;tot12[t]+=v12;tot13[t]+=v13;col+=3
                xd(ws_comp,row,col,gt12 if gt12 else None,bg=fill,b=True)
                xd(ws_comp,row,col+1,gt13 if gt13 else None,bg=fill,b=True)
                xdf(ws_comp,row,col+2,(gt13-gt12) if gt13-gt12!=0 else None,bg=fill);row+=1
            s12={t:sum(wk_prev.get(c,{t:0 for t in inv_types})[t] for c in gl) for t in inv_types}
            s13={t:sum(int(wk13_p.loc[c,t]) if c in wk13_p.index and t in wk13_p.columns else 0 for c in gl) for t in inv_types}
            ws_comp.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
            xh(ws_comp,row,1,f'{gn} Total',SUBTOT);col=3;sgt12=sgt13=0
            for t in inv_types:
                xh(ws_comp,row,col,s12[t] if s12[t] else None,SUBTOT)
                xh(ws_comp,row,col+1,s13[t] if s13[t] else None,SUBTOT)
                xdf(ws_comp,row,col+2,(s13[t]-s12[t]) if s13[t]-s12[t]!=0 else None,bg=SUBTOT,hmode=True)
                sgt12+=s12[t];sgt13+=s13[t];col+=3
            xh(ws_comp,row,col,sgt12,SUBTOT);xh(ws_comp,row,col+1,sgt13,SUBTOT)
            xdf(ws_comp,row,col+2,sgt13-sgt12,bg=SUBTOT,hmode=True);row+=1
        ws_comp.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
        xh(ws_comp,row,1,'Grand Total',GRAND);col=3;ggt12=ggt13=0
        for t in inv_types:
            xh(ws_comp,row,col,tot12[t] if tot12[t] else None,GRAND)
            xh(ws_comp,row,col+1,tot13[t] if tot13[t] else None,GRAND)
            xdf(ws_comp,row,col+2,(tot13[t]-tot12[t]) if tot13[t]-tot12[t]!=0 else None,bg=GRAND,hmode=True)
            ggt12+=tot12[t];ggt13+=tot13[t];col+=3
        xh(ws_comp,row,col,ggt12,GRAND);xh(ws_comp,row,col+1,ggt13,GRAND)
        xdf(ws_comp,row,col+2,ggt13-ggt12,bg=GRAND,hmode=True)
        ws_comp.column_dimensions['A'].width=12;ws_comp.column_dimensions['B'].width=36
        col=3
        for _ in inv_types:
            ws_comp.column_dimensions[get_column_letter(col)].width=11
            ws_comp.column_dimensions[get_column_letter(col+1)].width=11
            ws_comp.column_dimensions[get_column_letter(col+2)].width=10;col+=3
        for i in range(3):ws_comp.column_dimensions[get_column_letter(col+i)].width=12
        ws_comp.freeze_panes='C4'

    buf=io.BytesIO();wb.save(buf);buf.seek(0);return buf

# ── TLP Excel ──
def build_excel_tlp(df,wk_prev=None,wk_label='WK13'):
    inv_tlp=['TLP Irregulars','TLP Printed Excess','TLP sin clasificacion','TLP Blanks Excess','Wip']
    years=sorted(df['Year'].dropna().unique())
    all_clients=[c for c in TLP_ORDER if c in df['Customer Name'].unique()]+[c for c in df['Customer Name'].unique() if c not in TLP_ORDER]
    wb=Workbook();wb.remove(wb.active)

    # Resumen Total
    ws1=wb.create_sheet('Resumen Total')
    p=df.pivot_table(index='Customer Name',columns='Clasificacion',values='Quantity',aggfunc='sum',fill_value=0)
    for t in inv_tlp:
        if t not in p.columns:p[t]=0
    p['Grand Total']=p[inv_tlp].sum(axis=1)
    ws1.merge_cells(start_row=1,start_column=3,end_row=1,end_column=2+len(inv_tlp))
    ws1['C1']='Inventory Type';ws1['C1'].font=Font(bold=True,name='Calibri',size=10,color=NAVY)
    ws1['C1'].alignment=Alignment(horizontal='center')
    for ci,h in enumerate(['Cliente','Style CustomerName']+inv_tlp+['Grand Total'],1):xh(ws1,2,ci,h)
    for i,cli in enumerate(all_clients):
        if cli not in p.index:continue
        r=p.loc[cli];fill=NAVY_ALT if i%2==0 else None
        ws1.cell(row=3+i,column=1,value='').border=brd()
        xd(ws1,3+i,2,cli,bg=fill,ha='left')
        for ci,col in enumerate(inv_tlp+['Grand Total'],3):xd(ws1,3+i,ci,int(r[col]) if r[col]!=0 else None,bg=fill)
    last=3+len(all_clients);gt=p[inv_tlp+['Grand Total']].sum()
    ws1.merge_cells(start_row=last,start_column=1,end_row=last,end_column=2);xh(ws1,last,1,'Grand Total',GRAND)
    for ci,col in enumerate(inv_tlp+['Grand Total'],3):xh(ws1,last,ci,int(gt[col]) if gt[col]!=0 else None,GRAND)
    ws1.column_dimensions['A'].width=5;ws1.column_dimensions['B'].width=22
    for ci in range(3,3+len(inv_tlp)+2):ws1.column_dimensions[get_column_letter(ci)].width=22
    ws1.freeze_panes='B3'

    # Antigüedad
    ws_age=wb.create_sheet('Antigüedad')
    p2=df.pivot_table(index='Customer Name',columns='Year',values='Quantity',aggfunc='sum',fill_value=0)
    for yr in years:
        if yr not in p2.columns:p2[yr]=0
    p2=p2[[yr for yr in years]];p2['Grand Total']=p2.sum(axis=1)
    ncols=2+len(years)+1
    ws_age.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
    c=ws_age.cell(row=1,column=1,value='  Antigüedad del Inventario TLP')
    c.font=Font(bold=True,color=WHITE,name='Calibri',size=12)
    c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
    c.border=brd();ws_age.row_dimensions[1].height=22
    ws_age.merge_cells(start_row=2,start_column=3,end_row=2,end_column=2+len(years))
    xh(ws_age,2,1,'',NAVY_MID);xh(ws_age,2,2,'Customer Name',NAVY_MID);xh(ws_age,2,3,'Year of Creation',NAVY_MID)
    for ci in range(4,2+len(years)+1):
        c=ws_age.cell(row=2,column=ci);c.fill=PatternFill('solid',start_color=NAVY_MID);c.border=brd()
    xh(ws_age,2,2+len(years)+1,'Grand Total',NAVY_MID)
    xh(ws_age,3,1,'',NAVY_LIGHT);xh(ws_age,3,2,'',NAVY_LIGHT)
    for yi,yr in enumerate(years):xh(ws_age,3,3+yi,int(yr),NAVY_LIGHT)
    xh(ws_age,3,3+len(years),'Grand Total',NAVY_LIGHT)
    for i,cli in enumerate(all_clients):
        if cli not in p2.index:continue
        r=p2.loc[cli];rb=NAVY_ALT if i%2==0 else None
        ws_age.cell(row=4+i,column=1,value='').border=brd()
        xd(ws_age,4+i,2,cli,bg=rb,ha='left')
        for yi,yr in enumerate(years):xd(ws_age,4+i,3+yi,int(r[yr]) if r[yr]!=0 else None,bg=rb)
        xd(ws_age,4+i,3+len(years),int(r['Grand Total']),bg=NAVY_PALE,b=True)
    last2=4+len(all_clients);gt2=p2.sum()
    ws_age.merge_cells(start_row=last2,start_column=1,end_row=last2,end_column=2);xh(ws_age,last2,1,'Grand Total',GRAND)
    for yi,yr in enumerate(years):xh(ws_age,last2,3+yi,int(gt2[yr]) if gt2[yr]!=0 else None,GRAND)
    xh(ws_age,last2,3+len(years),int(gt2['Grand Total']),GRAND)
    ws_age.column_dimensions['A'].width=5;ws_age.column_dimensions['B'].width=22
    for ci in range(3,3+len(years)+1):ws_age.column_dimensions[get_column_letter(ci)].width=10
    ws_age.column_dimensions[get_column_letter(3+len(years))].width=13;ws_age.freeze_panes='C4'

    # Damage Severity
    ws_dmg=wb.create_sheet('Damage Severity')
    dd=df[df['Clasificacion']=='TLP Irregulars'].copy()
    dd['DS_Group']=dd['Damage Severity'].apply(lambda x:'0, 1' if x in [0,1] else '2+')
    pd3=dd.pivot_table(index='Customer Name',columns='DS_Group',values='Quantity',aggfunc='sum',fill_value=0)
    for c in ['0, 1','2+']:
        if c not in pd3.columns:pd3[c]=0
    pd3=pd3[['0, 1','2+']];pd3['Grand Total']=pd3.sum(axis=1)
    ordered=[c for c in TLP_ORDER if c in pd3.index]+[c for c in pd3.index if c not in TLP_ORDER]
    pd3=pd3.reindex(ordered)
    ws_dmg.merge_cells('A1:E1')
    c=ws_dmg.cell(row=1,column=1,value='  Damage Severity — TLP Irregulars Summary')
    c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
    c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
    c.border=brd();ws_dmg.row_dimensions[1].height=26
    ws_dmg.merge_cells('C2:D2')
    xh(ws_dmg,2,1,'Inventory Type',NAVY_MID);xh(ws_dmg,2,2,'Style CustomerName',NAVY_MID)
    xh(ws_dmg,2,3,'Damage Severity',NAVY_MID)
    ws_dmg.cell(row=2,column=4).fill=PatternFill('solid',start_color=NAVY_MID);ws_dmg.cell(row=2,column=4).border=brd()
    xh(ws_dmg,2,5,'Grand Total',NAVY_MID)
    xh(ws_dmg,3,1,'',NAVY_LIGHT);xh(ws_dmg,3,2,'',NAVY_LIGHT)
    xh(ws_dmg,3,3,'0, 1',NAVY_LIGHT);xh(ws_dmg,3,4,'2 +',NAVY_LIGHT);xh(ws_dmg,3,5,'Grand Total',NAVY_LIGHT)
    row=4;first=True
    for i,(cli,r) in enumerate(pd3.iterrows()):
        fill=NAVY_ALT if i%2==0 else None
        c1=ws_dmg.cell(row=row,column=1,value='TLP Irregulars' if first else '')
        c1.font=Font(bold=True,color=WHITE,name='Calibri',size=10)
        c1.fill=PatternFill('solid',start_color=ACCENT)
        c1.alignment=Alignment(horizontal='center',vertical='center');c1.border=brd();first=False
        xd(ws_dmg,row,2,cli,bg=fill,ha='left')
        xd(ws_dmg,row,3,int(r['0, 1']) if r['0, 1']!=0 else None,bg=fill)
        xd(ws_dmg,row,4,int(r['2+']) if r['2+']!=0 else None,bg=fill)
        xd(ws_dmg,row,5,int(r['Grand Total']),bg=fill,b=True);row+=1
    gt3=pd3.sum()
    ws_dmg.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2);xh(ws_dmg,row,1,'Grand Total',GRAND)
    xh(ws_dmg,row,3,int(gt3['0, 1']) if gt3['0, 1']!=0 else None,GRAND)
    xh(ws_dmg,row,4,int(gt3['2+']) if gt3['2+']!=0 else None,GRAND)
    xh(ws_dmg,row,5,int(gt3['Grand Total']),GRAND)
    for col,w in zip('ABCDE',[16,28,14,14,14]):ws_dmg.column_dimensions[col].width=w
    ws_dmg.freeze_panes='C4'

    # Program
    ws_prog=wb.create_sheet('Program')
    dp=df[df['Clasificacion']!='Wip'].copy()
    pp=dp.pivot_table(index='Clasificacion',columns='Program',values='Quantity',aggfunc='sum',fill_value=0)
    for c in ['BLANKS','PRINTED']:
        if c not in pp.columns:pp[c]=0
    pp=pp[['BLANKS','PRINTED']];pp['Grand Total']=pp.sum(axis=1)
    order2=['TLP Irregulars','TLP Printed Excess','TLP sin clasificacion','TLP Blanks Excess']
    pp=pp.reindex([c for c in order2 if c in pp.index]+[c for c in pp.index if c not in order2])
    ws_prog.merge_cells('A1:E1')
    c=ws_prog.cell(row=1,column=1,value='  Program Summary — BLANKS vs PRINTED  (Finished Goods)')
    c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
    c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
    c.border=brd();ws_prog.row_dimensions[1].height=26
    ws_prog.merge_cells('B2:C2')
    xh(ws_prog,2,1,'Inventory Type',NAVY_MID);xh(ws_prog,2,2,'Program',NAVY_MID)
    ws_prog.cell(row=2,column=3).fill=PatternFill('solid',start_color=NAVY_MID);ws_prog.cell(row=2,column=3).border=brd()
    xh(ws_prog,2,4,'Grand Total',NAVY_MID)
    xh(ws_prog,3,1,'',NAVY_LIGHT);xh(ws_prog,3,2,'BLANKS',NAVY_LIGHT)
    xh(ws_prog,3,3,'PRINTED',NAVY_LIGHT);xh(ws_prog,3,4,'Grand Total',NAVY_LIGHT)
    for i,(clas,r) in enumerate(pp.iterrows()):
        fill=NAVY_ALT if i%2==0 else None
        xd(ws_prog,4+i,1,clas,bg=fill,ha='left')
        xd(ws_prog,4+i,2,int(r['BLANKS']) if r['BLANKS']!=0 else None,bg=fill)
        xd(ws_prog,4+i,3,int(r['PRINTED']) if r['PRINTED']!=0 else None,bg=fill)
        xd(ws_prog,4+i,4,int(r['Grand Total']),bg=fill,b=True)
    last3=4+len(pp);gt4=pp.sum()
    xh(ws_prog,last3,1,'Grand Total',GRAND,ha='left')
    xh(ws_prog,last3,2,int(gt4['BLANKS']) if gt4['BLANKS']!=0 else None,GRAND)
    xh(ws_prog,last3,3,int(gt4['PRINTED']) if gt4['PRINTED']!=0 else None,GRAND)
    xh(ws_prog,last3,4,int(gt4['Grand Total']),GRAND)
    for col,w in zip('ABCD',[26,16,16,16]):ws_prog.column_dimensions[col].width=w
    ws_prog.freeze_panes='A4'

    # Comparativo TLP
    if wk_prev:
        inv_comp=['TLP Irregulars','TLP Printed Excess','TLP sin clasificacion','TLP Blanks Excess']
        wk13_p=df.pivot_table(index='Customer Name',columns='Clasificacion',values='Quantity',aggfunc='sum',fill_value=0)
        for t in inv_comp:
            if t not in wk13_p.columns:wk13_p[t]=0
        ws_comp=wb.create_sheet('Comparativo')
        ncols=2+len(inv_comp)*3+3
        ws_comp.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
        c=ws_comp.cell(row=1,column=1,value=f'  Inventory Comparison TLP — {wk_label} vs WK Anterior  (sin Wip)')
        c.font=Font(bold=True,color=WHITE,name='Calibri',size=13)
        c.fill=PatternFill('solid',start_color=NAVY);c.alignment=Alignment(horizontal='left',vertical='center')
        c.border=brd();ws_comp.row_dimensions[1].height=26
        xh(ws_comp,2,1,'',NAVY_MID);xh(ws_comp,2,2,'Customer Name',NAVY_MID)
        col=3
        for t in inv_comp:
            ws_comp.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+2)
            xh(ws_comp,2,col,t,NAVY_MID);col+=3
        ws_comp.merge_cells(start_row=2,start_column=col,end_row=2,end_column=col+2)
        xh(ws_comp,2,col,'Grand Total',NAVY)
        xh(ws_comp,3,1,'',NAVY_LIGHT);xh(ws_comp,3,2,'',NAVY_LIGHT)
        col=3
        for t in inv_comp:
            xh(ws_comp,3,col,'WK Ant.',NAVY_LIGHT);xh(ws_comp,3,col+1,wk_label,NAVY_LIGHT);xh(ws_comp,3,col+2,'Diff',NAVY_LIGHT);col+=3
        xh(ws_comp,3,col,'WK Ant.',NAVY);xh(ws_comp,3,col+1,wk_label,NAVY);xh(ws_comp,3,col+2,'Diff',NAVY)
        row=4;tot12={t:0 for t in inv_comp};tot13={t:0 for t in inv_comp}
        for i,cli in enumerate(all_clients):
            w12=wk_prev.get(cli,{t:0 for t in inv_comp})
            w13={t:int(wk13_p.loc[cli,t]) if cli in wk13_p.index and t in wk13_p.columns else 0 for t in inv_comp}
            fill=NAVY_ALT if i%2==0 else None
            ws_comp.cell(row=row,column=1,value='').border=brd()
            xd(ws_comp,row,2,cli,bg=fill,ha='left')
            col=3;gt12=gt13=0
            for t in inv_comp:
                v12=w12[t];v13=w13[t];diff=v13-v12
                xd(ws_comp,row,col,v12 if v12 else None,bg=fill)
                xd(ws_comp,row,col+1,v13 if v13 else None,bg=fill)
                xdf(ws_comp,row,col+2,diff if diff!=0 else None,bg=fill)
                gt12+=v12;gt13+=v13;tot12[t]+=v12;tot13[t]+=v13;col+=3
            xd(ws_comp,row,col,gt12 if gt12 else None,bg=fill,b=True)
            xd(ws_comp,row,col+1,gt13 if gt13 else None,bg=fill,b=True)
            xdf(ws_comp,row,col+2,(gt13-gt12) if gt13-gt12!=0 else None,bg=fill);row+=1
        ws_comp.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2)
        xh(ws_comp,row,1,'Grand Total',GRAND);col=3;ggt12=ggt13=0
        for t in inv_comp:
            xh(ws_comp,row,col,tot12[t] if tot12[t] else None,GRAND)
            xh(ws_comp,row,col+1,tot13[t] if tot13[t] else None,GRAND)
            xdf(ws_comp,row,col+2,(tot13[t]-tot12[t]) if tot13[t]-tot12[t]!=0 else None,bg=GRAND,hmode=True)
            ggt12+=tot12[t];ggt13+=tot13[t];col+=3
        xh(ws_comp,row,col,ggt12,GRAND);xh(ws_comp,row,col+1,ggt13,GRAND)
        xdf(ws_comp,row,col+2,ggt13-ggt12,bg=GRAND,hmode=True)
        ws_comp.column_dimensions['A'].width=5;ws_comp.column_dimensions['B'].width=22
        col=3
        for _ in inv_comp:
            ws_comp.column_dimensions[get_column_letter(col)].width=20
            ws_comp.column_dimensions[get_column_letter(col+1)].width=20
            ws_comp.column_dimensions[get_column_letter(col+2)].width=12;col+=3
        for i in range(3):ws_comp.column_dimensions[get_column_letter(col+i)].width=14
        ws_comp.freeze_panes='C4'

    buf=io.BytesIO();wb.save(buf);buf.seek(0);return buf

# ══════════════════════════════════════════
# UI
# ══════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📦 Inventory Hub")
    st.markdown("---")
    week_label=st.text_input("Semana actual",value="WK13",label_visibility="collapsed")
    st.markdown(f"**Semana: {week_label}**")




# ══════════════════════════════════════════
# UI — Executive Dashboard
# ══════════════════════════════════════════

CLAS_COLORS_HN = {
    'Regular':'#2d5a3d','VMI':'#3B6D11','Irregulares':'#BA7517',
    'Exceso':'#993C1D','Obsoleto':'#A32D2D'
}
CLAS_COLORS_TLP = {
    'TLP Irregulars':'#BA7517','TLP Blanks Excess':'#993C1D',
    'TLP Printed Excess':'#791F1F','TLP sin clasificacion':'#185FA5','Wip':'#444441'
}

def make_bar_chart(data_series, color_map, title):
    labels = list(data_series.index)
    values = [int(v) for v in data_series.values]
    colors = [color_map.get(l, '#2d5a3d') for l in labels]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation='h',
        marker_color=colors,
        text=[f'{v:,}' for v in values],
        textposition='outside', textfont=dict(size=10)
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13)),
        height=max(220, len(labels)*32+80),
        margin=dict(t=40,b=10,l=10,r=70),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)',
                   tickformat=',', showticklabels=False),
        yaxis=dict(autorange='reversed', tickfont=dict(size=11))
    )
    return fig

def make_donut(data_series, color_map, title):
    labels = list(data_series.index)
    values = [int(v) for v in data_series.values]
    colors = [color_map.get(l, '#639922') for l in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, marker_colors=colors,
        hole=0.5, textinfo='percent', textfont=dict(size=10),
        hovertemplate='%{label}<br>%{value:,}<extra></extra>'
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13)),
        height=280, margin=dict(t=40,b=10,l=10,r=10),
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(font=dict(size=10), orientation='v', x=1.02)
    )
    return fig

def make_trend(r_cur, r_prev, week_label, color_map, title):
    cur = r_cur.groupby('Clasificacion')['Quantity'].sum()
    labels = list(cur.index)
    fig = go.Figure()
    if r_prev is not None and 'Clasificacion' in r_prev.columns:
        prev = r_prev.groupby('Clasificacion')['Quantity'].sum()
        prev_vals = [int(prev.get(l, 0)) for l in labels]
        fig.add_trace(go.Bar(name='WK Ant.', x=labels, y=prev_vals,
                             marker_color='#C0DD97'))
    cur_vals = [int(cur[l]) for l in labels]
    fig.add_trace(go.Bar(name=week_label, x=labels, y=cur_vals,
                         marker_color=[color_map.get(l,'#2d5a3d') for l in labels]))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13)),
        barmode='group', height=280,
        margin=dict(t=40,b=40,l=10,r=10),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)', tickformat=','),
        xaxis=dict(tickfont=dict(size=10)),
        legend=dict(font=dict(size=10), orientation='h', y=1.15)
    )
    return fig

def kpi_card(label, value, sub='', color='#2d5a3d', sub_color='#4CAF50'):
    return f"""<div style="background:var(--color-background-primary);border:0.5px solid
var(--color-border-tertiary);border-radius:var(--border-radius-lg);padding:1rem 1.2rem;
border-top:3px solid {color};">
<div style="font-size:11px;font-weight:500;color:var(--color-text-secondary);text-transform:uppercase;
letter-spacing:.08em;margin-bottom:6px;">{label}</div>
<div style="font-size:1.8rem;font-weight:600;color:var(--color-text-primary);line-height:1.1;
font-family:var(--font-mono);">{value}</div>
<div style="font-size:11px;color:{sub_color};margin-top:4px;">{sub}</div></div>"""

# ── Sidebar ──
with st.sidebar:
    st.markdown("## Inventory Hub")
    week_label = st.text_input("Semana", value="WK13", key="week_input")
    st.markdown("---")
    st.markdown("**Archivos Honduras**")
    carton_file   = st.file_uploader("Carton Report HN (CSV)", type=['csv'], key='hn_carton')
    open_file     = st.file_uploader("Open Order (CSV)", type=['csv'], key='hn_open')
    prev_hn_file  = st.file_uploader("HN semana anterior (CSV)", type=['csv'], key='hn_prev')
    st.markdown("**Archivos TLP**")
    tlp_file      = st.file_uploader("Carton Report TLP (CSV)", type=['csv'], key='tlp_carton')
    prev_tlp_file = st.file_uploader("TLP semana anterior (CSV)", type=['csv'], key='tlp_prev')
    st.markdown("---")
    if st.button("Clasificar ambas bodegas", type="primary", use_container_width=True):
        if carton_file and open_file:
            with st.spinner("Clasificando Honduras..."):
                carton_file.seek(0); open_file.seek(0)
                r_hn, cut = classify_honduras(
                    pd.read_csv(carton_file, low_memory=False),
                    pd.read_csv(open_file, low_memory=False)
                )
                st.session_state['hn_r'] = r_hn
                st.session_state['hn_cut'] = cut
        if prev_hn_file:
            prev_hn_file.seek(0)
            st.session_state['hn_prev_df'] = pd.read_csv(prev_hn_file, low_memory=False)
        else:
            st.session_state['hn_prev_df'] = None
        if tlp_file:
            with st.spinner("Clasificando TLP..."):
                tlp_file.seek(0)
                st.session_state['tlp_r'] = classify_tlp(pd.read_csv(tlp_file, low_memory=False))
        if prev_tlp_file:
            prev_tlp_file.seek(0)
            st.session_state['tlp_prev_df'] = pd.read_csv(prev_tlp_file, low_memory=False)
        else:
            st.session_state['tlp_prev_df'] = None
        st.success("Listo!")

# ── Tabs ──
tab_dash, tab_hn, tab_tlp, tab_comp, tab_dl = st.tabs([
    "Dashboard", "Honduras", "TLP", "Comparativo", "Descargas"
])

r_hn   = st.session_state.get('hn_r')
r_tlp  = st.session_state.get('tlp_r')
cut_n, cut_u = st.session_state.get('hn_cut', (0,0))
prev_hn  = st.session_state.get('hn_prev_df')
prev_tlp = st.session_state.get('tlp_prev_df')

# ── Dashboard ──
with tab_dash:
    if r_hn is None and r_tlp is None:
        st.info("Carga los archivos en el panel izquierdo y presiona Clasificar.")
    else:
        if r_hn is not None:
            st.markdown("### Honduras")
            tot_hn = int(r_hn['Quantity'].sum())
            fg_hn  = int(r_hn[r_hn['Type']=='Finished Goods']['Quantity'].sum())
            wip_hn = int(r_hn[r_hn['Type']=='Wip']['Quantity'].sum())
            prev_tot_hn = int(pd.to_numeric(prev_hn['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0).sum()) if prev_hn is not None and 'Quantity' in prev_hn.columns else None
            diff_hn = f"+{tot_hn-prev_tot_hn:,} vs WK ant." if prev_tot_hn else "Sin comparativo"
            c1,c2,c3,c4,c5 = st.columns(5)
            with c1: st.markdown(kpi_card("Total", f"{tot_hn:,}", diff_hn), unsafe_allow_html=True)
            with c2: st.markdown(kpi_card("Finished Goods", f"{fg_hn:,}", f"{fg_hn/tot_hn*100:.1f}%"), unsafe_allow_html=True)
            with c3: st.markdown(kpi_card("Wip", f"{wip_hn:,}", f"{wip_hn/tot_hn*100:.1f}%"), unsafe_allow_html=True)
            with c4: st.markdown(kpi_card("Cut excluidas", f"{cut_n:,}", f"{cut_u:,} uds", '#A32D2D','#E24B4A'), unsafe_allow_html=True)
            with c5: st.markdown(kpi_card("Tipos", f"{r_hn['Clasificacion'].nunique()}", "clasificaciones",'#185FA5','#378ADD'), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            ca,cb,cc = st.columns(3)
            with ca:
                cs = r_hn.groupby('Clasificacion')['Quantity'].sum().sort_values(ascending=False)
                st.plotly_chart(make_donut(cs, CLAS_COLORS_HN, "Clasificacion HN"), use_container_width=True)
            with cb:
                cli = r_hn.groupby('Customer Name')['Quantity'].sum().sort_values(ascending=True)
                st.plotly_chart(make_bar_chart(cli, {}, "Clientes HN (codigo)"), use_container_width=True)
            with cc:
                if prev_hn is not None and 'Clasificacion' in prev_hn.columns:
                    prev_hn2 = prev_hn.copy()
                    prev_hn2['Quantity'] = pd.to_numeric(prev_hn2['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    st.plotly_chart(make_trend(r_hn, prev_hn2, week_label, CLAS_COLORS_HN, "Tendencia HN"), use_container_width=True)
                else:
                    st.plotly_chart(make_trend(r_hn, None, week_label, CLAS_COLORS_HN, "Clasificacion HN"), use_container_width=True)

        st.markdown("---")

        if r_tlp is not None:
            st.markdown("### TLP")
            tot_tlp = int(r_tlp['Quantity'].sum())
            fg_tlp  = int(r_tlp[r_tlp['Clasificacion']!='Wip']['Quantity'].sum())
            wip_tlp = int(r_tlp[r_tlp['Clasificacion']=='Wip']['Quantity'].sum())
            prev_tot_tlp = int(pd.to_numeric(prev_tlp['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0).sum()) if prev_tlp is not None and 'Quantity' in prev_tlp.columns else None
            diff_tlp = f"+{tot_tlp-prev_tot_tlp:,} vs WK ant." if prev_tot_tlp else "Sin comparativo"
            c1,c2,c3,c4 = st.columns(4)
            with c1: st.markdown(kpi_card("Total", f"{tot_tlp:,}", diff_tlp,'#1a3a5c','#378ADD'), unsafe_allow_html=True)
            with c2: st.markdown(kpi_card("Finished Goods", f"{fg_tlp:,}", f"{fg_tlp/tot_tlp*100:.1f}%",'#1a3a5c','#378ADD'), unsafe_allow_html=True)
            with c3: st.markdown(kpi_card("Wip", f"{wip_tlp:,}", f"{wip_tlp/tot_tlp*100:.1f}%",'#444441','#888780'), unsafe_allow_html=True)
            with c4: st.markdown(kpi_card("Tipos", f"{r_tlp['Clasificacion'].nunique()}", "clasificaciones",'#185FA5','#378ADD'), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            ca,cb,cc = st.columns(3)
            with ca:
                cs2 = r_tlp.groupby('Clasificacion')['Quantity'].sum().sort_values(ascending=False)
                st.plotly_chart(make_donut(cs2, CLAS_COLORS_TLP, "Clasificacion TLP"), use_container_width=True)
            with cb:
                cli2 = r_tlp.groupby('Customer Name')['Quantity'].sum().sort_values(ascending=True)
                st.plotly_chart(make_bar_chart(cli2, {}, "Clientes TLP (codigo)"), use_container_width=True)
            with cc:
                if prev_tlp is not None and 'Clasificacion' in prev_tlp.columns:
                    prev_tlp2 = prev_tlp.copy()
                    prev_tlp2['Quantity'] = pd.to_numeric(prev_tlp2['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    st.plotly_chart(make_trend(r_tlp, prev_tlp2, week_label, CLAS_COLORS_TLP, "Tendencia TLP"), use_container_width=True)
                else:
                    st.plotly_chart(make_trend(r_tlp, None, week_label, CLAS_COLORS_TLP, "Clasificacion TLP"), use_container_width=True)

# ── Honduras tab ──
with tab_hn:
    if r_hn is None:
        st.info("Carga el Carton Report y Open Order en el panel izquierdo.")
    else:
        st.markdown(f"### Honduras — {week_label}")
        sm = r_hn.groupby('Clasificacion').agg(Cajas=('Quantity','count'),Unidades=('Quantity','sum')).reset_index().sort_values('Unidades',ascending=False)
        sm['Unidades'] = sm['Unidades'].astype(int)
        sm['% Total'] = (sm['Unidades']/sm['Unidades'].sum()*100).round(1).astype(str)+'%'
        st.dataframe(sm, use_container_width=True, hide_index=True)
        st.markdown("#### Por cliente — mayor a menor")
        cli_df = r_hn.groupby('Customer Name')['Quantity'].sum().sort_values(ascending=False).reset_index()
        cli_df.columns=['Codigo Cliente','Unidades']
        cli_df['Unidades'] = cli_df['Unidades'].astype(int)
        cli_df['% Total'] = (cli_df['Unidades']/cli_df['Unidades'].sum()*100).round(1).astype(str)+'%'
        st.dataframe(cli_df, use_container_width=True, hide_index=True)
        st.markdown("#### Por clasificacion y cliente")
        cross = r_hn.groupby(['Customer Name','Clasificacion'])['Quantity'].sum().reset_index()
        cross.columns=['Codigo','Clasificacion','Unidades']
        cross['Unidades'] = cross['Unidades'].astype(int)
        cross = cross.sort_values('Unidades',ascending=False)
        st.dataframe(cross, use_container_width=True, hide_index=True)

# ── TLP tab ──
with tab_tlp:
    if r_tlp is None:
        st.info("Carga el Carton Report TLP en el panel izquierdo.")
    else:
        st.markdown(f"### TLP — {week_label}")
        sm2 = r_tlp.groupby('Clasificacion').agg(Cajas=('Quantity','count'),Unidades=('Quantity','sum')).reset_index().sort_values('Unidades',ascending=False)
        sm2['Unidades'] = sm2['Unidades'].astype(int)
        sm2['% Total'] = (sm2['Unidades']/sm2['Unidades'].sum()*100).round(1).astype(str)+'%'
        st.dataframe(sm2, use_container_width=True, hide_index=True)
        st.markdown("#### Por cliente — mayor a menor")
        cli_df2 = r_tlp.groupby('Customer Name')['Quantity'].sum().sort_values(ascending=False).reset_index()
        cli_df2.columns=['Codigo Cliente','Unidades']
        cli_df2['Unidades'] = cli_df2['Unidades'].astype(int)
        cli_df2['% Total'] = (cli_df2['Unidades']/cli_df2['Unidades'].sum()*100).round(1).astype(str)+'%'
        st.dataframe(cli_df2, use_container_width=True, hide_index=True)

# ── Comparativo tab ──
with tab_comp:
    if r_hn is None and r_tlp is None:
        st.info("Clasifica primero los inventarios.")
    else:
        if r_hn is not None and prev_hn is not None and 'Clasificacion' in prev_hn.columns:
            st.markdown(f"### Honduras — {week_label} vs WK anterior")
            prev_hn3 = prev_hn.copy()
            prev_hn3['Quantity'] = pd.to_numeric(prev_hn3['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            cur_s  = r_hn.groupby('Clasificacion')['Quantity'].sum()
            prev_s = prev_hn3.groupby('Clasificacion')['Quantity'].sum()
            all_c  = sorted(set(list(cur_s.index)+list(prev_s.index)))
            comp_hn = pd.DataFrame({'Clasificacion':all_c,'WK Anterior':[int(prev_s.get(c,0)) for c in all_c],week_label:[int(cur_s.get(c,0)) for c in all_c]})
            comp_hn['Diferencia'] = comp_hn[week_label] - comp_hn['WK Anterior']
            comp_hn = comp_hn.sort_values(week_label, ascending=False)
            st.dataframe(comp_hn, use_container_width=True, hide_index=True)
            st.plotly_chart(make_trend(r_hn, prev_hn3, week_label, CLAS_COLORS_HN, "Tendencia Honduras"), use_container_width=True)
        elif r_hn is not None:
            st.warning("Carga el inventario HN de la semana anterior para ver el comparativo.")
        st.markdown("---")
        if r_tlp is not None and prev_tlp is not None and 'Clasificacion' in prev_tlp.columns:
            st.markdown(f"### TLP — {week_label} vs WK anterior")
            prev_tlp3 = prev_tlp.copy()
            prev_tlp3['Quantity'] = pd.to_numeric(prev_tlp3['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            cur_s2  = r_tlp.groupby('Clasificacion')['Quantity'].sum()
            prev_s2 = prev_tlp3.groupby('Clasificacion')['Quantity'].sum()
            all_c2  = sorted(set(list(cur_s2.index)+list(prev_s2.index)))
            comp_tlp = pd.DataFrame({'Clasificacion':all_c2,'WK Anterior':[int(prev_s2.get(c,0)) for c in all_c2],week_label:[int(cur_s2.get(c,0)) for c in all_c2]})
            comp_tlp['Diferencia'] = comp_tlp[week_label] - comp_tlp['WK Anterior']
            comp_tlp = comp_tlp.sort_values(week_label, ascending=False)
            st.dataframe(comp_tlp, use_container_width=True, hide_index=True)
            st.plotly_chart(make_trend(r_tlp, prev_tlp3, week_label, CLAS_COLORS_TLP, "Tendencia TLP"), use_container_width=True)
        elif r_tlp is not None:
            st.warning("Carga el inventario TLP de la semana anterior para ver el comparativo.")

# ── Descargas tab ──
with tab_dl:
    st.markdown("### Descargar archivos")
    if r_hn is not None:
        st.markdown("#### Honduras")
        wk_prev_hn = None
        if prev_hn is not None and 'Clasificacion' in prev_hn.columns and 'Customer Name' in prev_hn.columns:
            prev_hn4 = prev_hn.copy()
            prev_hn4['Quantity'] = pd.to_numeric(prev_hn4['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            inv_types = ['Regulars','VMI','Excess','Irregulars','Obsolete','Liability']
            prev_hn4['Clas_Col'] = prev_hn4['Clasificacion'].map(CMAP_HN)
            pfg = prev_hn4[prev_hn4['Type']=='Finished Goods'] if 'Type' in prev_hn4.columns else prev_hn4
            pp = pfg.pivot_table(index='Customer Name',columns='Clas_Col',values='Quantity',aggfunc='sum',fill_value=0)
            wk_prev_hn = {cli:{t:int(pp.loc[cli,t]) if cli in pp.index and t in pp.columns else 0 for t in inv_types} for cli in ACTIVOS+INACTIVOS}
        with st.spinner("Generando Excel Honduras..."):
            buf_hn = build_excel_hn(r_hn, wk_prev_hn, week_label)
        c1,c2 = st.columns(2)
        with c1:
            st.download_button(f"Excel Honduras {week_label}", data=buf_hn,
                file_name=f"inventario_honduras_{week_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            b2=io.StringIO(); r_hn.to_csv(b2,index=False)
            st.download_button("CSV Honduras completo",data=b2.getvalue(),
                file_name=f"data_honduras_{week_label}.csv",mime="text/csv",use_container_width=True)

    if r_tlp is not None:
        st.markdown("#### TLP")
        wk_prev_tlp = None
        if prev_tlp is not None and 'Clasificacion' in prev_tlp.columns and 'Customer Name' in prev_tlp.columns:
            prev_tlp4 = prev_tlp.copy()
            prev_tlp4['Quantity'] = pd.to_numeric(prev_tlp4['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
            inv_comp = ['TLP Irregulars','TLP Printed Excess','TLP sin clasificacion','TLP Blanks Excess']
            pp2 = prev_tlp4.pivot_table(index='Customer Name',columns='Clasificacion',values='Quantity',aggfunc='sum',fill_value=0)
            all_tlp_c = list(set(TLP_ORDER+list(pp2.index)))
            wk_prev_tlp = {cli:{t:int(pp2.loc[cli,t]) if cli in pp2.index and t in pp2.columns else 0 for t in inv_comp} for cli in all_tlp_c}
        with st.spinner("Generando Excel TLP..."):
            buf_tlp = build_excel_tlp(r_tlp, wk_prev_tlp, week_label)
        c1,c2 = st.columns(2)
        with c1:
            st.download_button(f"Excel TLP {week_label}", data=buf_tlp,
                file_name=f"inventario_TLP_{week_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c2:
            b4=io.StringIO(); r_tlp.to_csv(b4,index=False)
            st.download_button("CSV TLP completo",data=b4.getvalue(),
                file_name=f"data_TLP_{week_label}.csv",mime="text/csv",use_container_width=True)

    if r_hn is None and r_tlp is None:
        st.info("Clasifica primero los inventarios para poder descargar.")
