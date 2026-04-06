import streamlit as st
import pandas as pd
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
    st.markdown("---")
    st.markdown("**Reglas Honduras:**")
    rules_hn=["Excluir Cut","Picked incluido","Is Second/Third → Irregulares","Locker Stock → VMI",
              "Sin VMI: Duluth, Tiltworks, Vortex","PO Open Order → Regular","Box Tag → Exceso/Obsoleto",
              "Create Date → Obsoleto/Exceso","FG vs Wip por Box Status"]
    for i,r in enumerate(rules_hn,1):
        st.markdown(f"<div style='font-size:.78rem;padding:2px 0;color:#a8c4e0;'><b style='color:#1B6CA8;'>{i}.</b> {r}</div>",unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**Reglas TLP:**")
    rules_tlp=["Is Second/Third → TLP Irregulars","Blanks Excess → TLP Blanks Excess",
               "Printed Excess → TLP Printed Excess","Packed/Picked → TLP sin clasificacion","Resto → Wip"]
    for i,r in enumerate(rules_tlp,1):
        st.markdown(f"<div style='font-size:.78rem;padding:2px 0;color:#a8c4e0;'><b style='color:#1B6CA8;'>{i}.</b> {r}</div>",unsafe_allow_html=True)



# ══════════════════════════════════════════
# UI — Full App
# ══════════════════════════════════════════

import plotly.express as px
import plotly.graph_objects as go

import pandas as pd

def parse_prev_hn(df_raw):
    try:
        clas_map = {'Regulars':'Regular','VMI':'VMI','excess':'Exceso',
                    'Irregulars':'Irregulares','Obsolete':'Obsoleto','Liability':'Liability'}
        result = {}
        for i in range(38, min(52, len(df_raw))):
            row = df_raw.iloc[i]
            for col_idx in [1, 5, 7]:
                if col_idx >= len(row): continue
                clas_raw = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ''
                if col_idx+1 >= len(row): continue
                val_raw = str(row.iloc[col_idx+1]).strip() if pd.notna(row.iloc[col_idx+1]) else ''
                if clas_raw in clas_map and val_raw not in ['nan','-','','NaN']:
                    try:
                        val = int(float(val_raw.replace(',','')))
                        result[clas_map[clas_raw]] = result.get(clas_map[clas_raw], 0) + val
                    except: pass
        return result if result else None
    except:
        return None

def parse_prev_tlp(df_raw):
    try:
        clas_map = {'TLP irregulars':'TLP Irregulars','TLP printed excess':'TLP Printed Excess',
                    'TLP sin clasificacion':'TLP sin clasificacion','TLP Blanks excess':'TLP Blanks Excess'}
        result = {}
        headers = [str(h).strip() for h in df_raw.iloc[1].tolist()]
        for i in range(2, len(df_raw)):
            row = df_raw.iloc[i]
            client = str(row.iloc[0]).strip()
            if client in ['','nan','Grand Total','NaN']: continue
            for j, h in enumerate(headers[1:], 1):
                if h in clas_map and j < len(row):
                    val_raw = str(row.iloc[j]).strip()
                    if val_raw not in ['nan','-','','NaN']:
                        try:
                            val = int(float(val_raw.replace(',','')))
                            result[clas_map[h]] = result.get(clas_map[h], 0) + val
                        except: pass
        return result if result else None
    except:
        return None

def is_pivot_format(df):
    cols = [str(c).lower() for c in df.columns.tolist()]
    return 'clasificacion' not in cols and 'customer name' not in cols


CLAS_COLORS_HN = {
    'Regular':'#2d5a3d','VMI':'#3B6D11','Irregulares':'#BA7517',
    'Exceso':'#993C1D','Obsoleto':'#A32D2D','Regular Wip':'#5F5E5A'
}
CLAS_COLORS_TLP = {
    'TLP Irregulars':'#BA7517','TLP Blanks Excess':'#993C1D',
    'TLP Printed Excess':'#791F1F','TLP sin clasificacion':'#185FA5','Wip':'#444441'
}
HN_FG_CLAS  = ['Regular','VMI','Irregulares','Exceso','Obsoleto']
HN_WIP_CLAS = ['Regular']
TLP_FG_CLAS = ['TLP Irregulars','TLP Blanks Excess','TLP Printed Excess','TLP sin clasificacion']
TLP_WIP_CLAS= ['Wip']

def fmt(n): return f"{int(n):,}"

def filter_df(df, view, fg_clas, wip_clas):
    if view == 'all':
        if 'Type' in df.columns:
            return df[df['Type'].isin(['Finished Goods','Wip'])]
        return df[df['Clasificacion'].isin(fg_clas + wip_clas)]
    if view == 'fg':
        if 'Type' in df.columns:
            return df[df['Type']=='Finished Goods']
        return df[df['Clasificacion'].isin(fg_clas)]
    if view == 'wip':
        if 'Type' in df.columns:
            return df[df['Type']=='Wip']
        return df[df['Clasificacion'].isin(wip_clas)]
    return df

def make_donut(data_series, color_map, title):
    labels = list(data_series.index)
    values = [int(v) for v in data_series.values]
    colors = [color_map.get(l,'#888780') for l in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, marker_colors=colors,
        hole=0.52, textinfo='percent', textfont=dict(size=13),
        hovertemplate='%{label}<br>%{value:,}<extra></extra>'
    ))
    fig.update_layout(
        height=380, margin=dict(t=20,b=20,l=20,r=20),
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(font=dict(size=13), orientation='v', x=1.02),
        annotations=[dict(text=title, x=0.5, y=0.5, font=dict(size=13,color='#374151'),
                         showarrow=False, xref='paper', yref='paper')]
    )
    return fig

def color_detail(data_series, color_map):
    total = int(data_series.sum())
    html = ""
    for clas, val in data_series.items():
        color = color_map.get(clas, '#888780')
        pct = val/total*100 if total > 0 else 0
        html += f"""<div style="display:flex;align-items:center;gap:12px;padding:9px 0;
        border-bottom:0.5px solid var(--color-border-tertiary);">
        <div style="width:14px;height:14px;border-radius:50%;background:{color};flex-shrink:0;"></div>
        <div style="flex:1;font-size:15px;color:var(--color-text-primary);">{clas}</div>
        <div style="font-size:15px;font-weight:600;color:var(--color-text-primary);">{int(val):,}</div>
        <div style="font-size:13px;color:var(--color-text-secondary);width:48px;text-align:right;">{pct:.1f}%</div>
        </div>"""
    return html

def kpi_card(label, value, sub='', color='#2d5a3d', sub_color='#4CAF50'):
    return f"""<div style="background:var(--color-background-primary);border:0.5px solid
var(--color-border-tertiary);border-radius:var(--border-radius-lg);padding:1rem 1.2rem;
border-top:3px solid {color};">
<div style="font-size:11px;font-weight:500;color:var(--color-text-secondary);text-transform:uppercase;
letter-spacing:.08em;margin-bottom:6px;">{label}</div>
<div style="font-size:1.7rem;font-weight:600;color:var(--color-text-primary);line-height:1.1;
font-family:var(--font-mono);">{value}</div>
<div style="font-size:11px;color:{sub_color};margin-top:4px;">{sub}</div></div>"""

def build_alerts(df_cur, df_prev, clas_colors):
    df_cur = df_cur.copy()
    df_cur['Quantity'] = pd.to_numeric(df_cur['Quantity'], errors='coerce').fillna(0)
    cur_cli = df_cur.groupby('Customer Name')['Quantity'].sum()

    html_summary = ""
    if df_prev is not None and ('Clasificacion' in df_prev.columns or is_pivot_format(df_prev)):
        if is_pivot_format(df_prev):
            return html_summary
        df_prev = df_prev.copy()
        df_prev['Quantity'] = pd.to_numeric(df_prev['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
        prev_cli = df_prev.groupby('Customer Name')['Quantity'].sum() if 'Customer Name' in df_prev.columns else pd.Series()

        all_cli = set(list(cur_cli.index)+list(prev_cli.index))
        diffs = {c: int(cur_cli.get(c,0)-prev_cli.get(c,0)) for c in all_cli}

        top_up  = sorted(diffs.items(), key=lambda x: x[1], reverse=True)[:5]
        top_dn  = sorted(diffs.items(), key=lambda x: x[1])[:5]
        top_pct = sorted([(c,d/prev_cli.get(c,1)*100) for c,d in diffs.items() if prev_cli.get(c,0)>0],
                         key=lambda x: x[1], reverse=True)[:5]

        cur_clas = df_cur.groupby('Clasificacion')['Quantity'].sum()
        prev_clas= df_prev.groupby('Clasificacion')['Quantity'].sum() if 'Clasificacion' in df_prev.columns else pd.Series()
        clas_diff= {c: int(cur_clas.get(c,0)-prev_clas.get(c,0)) for c in set(list(cur_clas.index)+list(prev_clas.index))}
        top_clas = sorted(clas_diff.items(), key=lambda x: x[1], reverse=True)[:5]

        # Summary text
        total_diff = int(cur_cli.sum() - prev_cli.sum())
        sign = '+' if total_diff >= 0 else ''
        biggest_up = top_up[0] if top_up else None
        biggest_clas = max(clas_diff.items(), key=lambda x: x[1]) if clas_diff else None

        summary = f"El inventario total varió <b>{sign}{total_diff:,} uds</b> esta semana. "
        if biggest_up:
            summary += f"El cliente <b>{biggest_up[0]}</b> es el mayor contribuyente al cambio con <b>{'+' if biggest_up[1]>=0 else ''}{biggest_up[1]:,} uds</b>. "
        if biggest_clas and biggest_clas[1] != 0:
            pct_clas = biggest_clas[1]/prev_clas.get(biggest_clas[0],1)*100
            summary += f"La clasificación <b>{biggest_clas[0]}</b> {'creció' if biggest_clas[1]>0 else 'bajó'} un <b>{pct_clas:+.0f}%</b> — requiere atención."

        html_summary = f"""<div style="background:var(--color-background-secondary);border:0.5px solid
        var(--color-border-tertiary);border-radius:var(--border-radius-lg);padding:14px 16px;margin-bottom:16px;">
        <div style="font-size:12px;font-weight:500;color:var(--color-text-primary);margin-bottom:8px;">Análisis automático</div>
        <div style="font-size:12px;color:var(--color-text-secondary);line-height:1.7;">{summary}</div>
        </div>"""

        def alert_rows(items, rank_color, val_color, show_pct=True, is_pct=False):
            max_val = abs(items[0][1]) if items else 1
            rows = ""
            for i,(code,val) in enumerate(items,1):
                bar_w = int(abs(val)/max_val*100) if max_val else 0
                sign = '+' if val >= 0 else ''
                val_str = f"{sign}{val:.0f}%" if is_pct else f"{sign}{int(val):,}"
                rows += f"""<div style="display:flex;align-items:center;gap:8px;padding:6px 0;
                border-bottom:0.5px solid var(--color-border-tertiary);">
                <div style="width:20px;height:20px;border-radius:50%;background:{rank_color}22;
                display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:600;
                color:{rank_color};flex-shrink:0;">{i}</div>
                <div style="font-size:12px;font-weight:600;color:var(--color-text-primary);min-width:36px;">{code}</div>
                <div style="flex:1;background:var(--color-background-secondary);border-radius:2px;height:5px;">
                  <div style="width:{bar_w}%;height:5px;border-radius:2px;background:{rank_color};opacity:.7;"></div>
                </div>
                <div style="font-size:11px;font-weight:500;color:{val_color};min-width:64px;text-align:right;">{val_str}</div>
                </div>"""
            return rows

        def alert_card(title, sub, icon, bg, rank_color, val_color, rows_html):
            return f"""<div style="background:var(--color-background-primary);border:0.5px solid
            var(--color-border-tertiary);border-radius:var(--border-radius-lg);padding:14px;">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
              <div style="width:28px;height:28px;border-radius:8px;background:{bg};display:flex;
              align-items:center;justify-content:center;font-size:13px;flex-shrink:0;">{icon}</div>
              <div>
                <div style="font-size:12px;font-weight:500;color:var(--color-text-primary);">{title}</div>
                <div style="font-size:10px;color:var(--color-text-secondary);">{sub}</div>
              </div>
            </div>
            {rows_html}
            </div>"""

        c1,c2 = st.columns(2)
        with c1:
            st.markdown(alert_card("Mayor incremento","Top 5 clientes que más subieron","▲","#DCFCE7","#1B5E20","#166534",
                alert_rows(top_up,'#1B5E20','#166534')), unsafe_allow_html=True)
        with c2:
            st.markdown(alert_card("Mayor reducción","Top 5 clientes que más bajaron","▼","#FEE2E2","#A32D2D","#991B1B",
                alert_rows(top_dn,'#A32D2D','#991B1B')), unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)
        c3,c4 = st.columns(2)
        with c3:
            st.markdown(alert_card("Mayor % de crecimiento","Clientes que más crecieron proporcionalmente","!","#FEF9C3","#BA7517","#854D0E",
                alert_rows(top_pct,'#BA7517','#854D0E',is_pct=True)), unsafe_allow_html=True)
        with c4:
            rows_clas = ""
            max_v = abs(top_clas[0][1]) if top_clas else 1
            for i,(clas,val) in enumerate(top_clas,1):
                bar_w = int(abs(val)/max_v*100) if max_v else 0
                color = clas_colors.get(clas,'#7C3AED')
                sign = '+' if val>=0 else ''
                rows_clas += f"""<div style="display:flex;align-items:center;gap:8px;padding:6px 0;
                border-bottom:0.5px solid var(--color-border-tertiary);">
                <div style="width:20px;height:20px;border-radius:50%;background:#EDE9FE;
                display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:600;
                color:#5B21B6;flex-shrink:0;">{i}</div>
                <div style="font-size:11px;font-weight:600;color:var(--color-text-primary);min-width:80px;">{clas}</div>
                <div style="flex:1;background:var(--color-background-secondary);border-radius:2px;height:5px;">
                  <div style="width:{bar_w}%;height:5px;border-radius:2px;background:{color};opacity:.7;"></div>
                </div>
                <div style="font-size:11px;font-weight:500;color:#5B21B6;min-width:64px;text-align:right;">{sign}{val:,}</div>
                </div>"""
            st.markdown(alert_card("Clasificaciones críticas","Las que más cambiaron esta semana","★","#EDE9FE","#7C3AED","#5B21B6",
                rows_clas), unsafe_allow_html=True)
    else:
        st.info("Carga el inventario de la semana anterior para ver las alertas.")

    return html_summary

def render_client_table(df, all_clas, color_map, theme_color, theme_header_bg):
    df = df.copy()
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
    pivot = df.pivot_table(index='Customer Name', columns='Clasificacion', values='Quantity', aggfunc='sum', fill_value=0)
    cajas = df.groupby('Customer Name').size()
    for c in all_clas:
        if c not in pivot.columns: pivot[c] = 0
    pivot['_total'] = pivot[all_clas].sum(axis=1)
    pivot = pivot.sort_values('_total', ascending=False)
    total_uds = int(pivot['_total'].sum())

    th = f"padding:9px 10px;text-align:right;font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.04em;color:{theme_color};background:{theme_header_bg};"
    th_l = th.replace("text-align:right","text-align:left")
    td_s = "border-bottom:0.5px solid var(--color-border-tertiary);padding:8px 10px;font-size:12px;"

    headers = f'<th style="{th_l}">Código</th><th style="{th}">Cajas</th><th style="{th}">Total Uds</th><th style="{th}">% Total</th>'
    for c in all_clas:
        dot = color_map.get(c,'#888')
        cname = c.replace('TLP ','')
        headers += f'<th style="{th}"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{dot};margin-right:4px;vertical-align:middle;"></span>{cname}</th>'

    rows_html = ""
    for i,(cli,row) in enumerate(pivot.iterrows()):
        bg = "background:var(--color-background-secondary);" if i%2==0 else "background:var(--color-background-primary);"
        pct = row['_total']/total_uds*100 if total_uds else 0
        rows_html += f'<tr style="{bg}">'
        rows_html += f'<td style="{td_s}text-align:left;font-weight:600;">{cli}</td>'
        rows_html += f'<td style="{td_s}text-align:right;">{int(cajas.get(cli,0)):,}</td>'
        rows_html += f'<td style="{td_s}text-align:right;font-weight:600;">{int(row["_total"]):,}</td>'
        rows_html += f'<td style="{td_s}text-align:right;color:var(--color-text-secondary);">{pct:.1f}%</td>'
        for c in all_clas:
            val = int(row.get(c,0))
            rows_html += f'<td style="{td_s}text-align:right;">{val:,}</td>' if val else f'<td style="{td_s}text-align:right;color:var(--color-text-secondary);">—</td>'
        rows_html += '</tr>'

    tot_td = f"padding:9px 10px;font-size:12px;font-weight:600;color:{theme_color};background:{theme_header_bg};border-top:2px solid var(--color-border-secondary);"
    rows_html += f'<tr><td style="{tot_td}text-align:left;">Total</td>'
    rows_html += f'<td style="{tot_td}text-align:right;">{int(cajas.sum()):,}</td>'
    rows_html += f'<td style="{tot_td}text-align:right;">{total_uds:,}</td>'
    rows_html += f'<td style="{tot_td}text-align:right;">100%</td>'
    for c in all_clas:
        t = int(pivot[c].sum())
        rows_html += f'<td style="{tot_td}text-align:right;">{t:,}</td>' if t else f'<td style="{tot_td}text-align:right;">—</td>'
    rows_html += '</tr>'

    st.markdown(
        f'<div style="overflow-x:auto;border:0.5px solid var(--color-border-tertiary);border-radius:var(--border-radius-lg);">'
        f'<table style="width:100%;border-collapse:collapse;font-family:var(--font-sans);white-space:nowrap;">'
        f'<thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table></div>',
        unsafe_allow_html=True
    )


def render_line_chart(weeks_data, weeks_labels, color, bodega):
    totals = [sum(d.values()) for d in weeks_data]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=weeks_labels, y=totals, mode='lines+markers',
        line=dict(color=color, width=2.5),
        marker=dict(size=8, color=color),
        fill='tozeroy', fillcolor='rgba('+str(int(color[1:3],16))+','+str(int(color[3:5],16))+','+str(int(color[5:7],16))+',0.1)',
        hovertemplate='%{x}: %{y:,} uds<extra></extra>'
    ))
    fig.update_layout(
        height=200, margin=dict(t=10,b=10,l=10,r=10),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(tickfont=dict(size=10), showgrid=False),
        yaxis=dict(tickfont=dict(size=10), tickformat=',', gridcolor='rgba(0,0,0,0.05)'),
        showlegend=False
    )
    return fig, totals

def wk_badges_html(totals, labels, color):
    html = '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">'
    for i,(t,l) in enumerate(zip(totals,labels)):
        if i == 0:
            diff_html = '<div style="font-size:10px;color:var(--color-text-secondary);">inicio</div>'
        else:
            diff = t - totals[i-1]
            pct = diff/totals[i-1]*100 if totals[i-1] else 0
            sign = '+' if diff >= 0 else ''
            col = '#166534' if diff >= 0 else '#991B1B'
            arr = '▲' if diff >= 0 else '▼'
            diff_html = f'<div style="font-size:10px;color:{col};font-weight:500;">{arr} {sign}{diff:,}<br><span style="font-size:9px;">{sign}{pct:.1f}%</span></div>'
        is_last = i == len(totals)-1
        border = f'border:1.5px solid {color};' if is_last else 'border:0.5px solid var(--color-border-tertiary);'
        html += f"""<div style="background:var(--color-background-primary);{border}
        border-radius:10px;padding:8px 12px;text-align:center;min-width:72px;">
        <div style="font-size:10px;color:var(--color-text-secondary);font-weight:500;margin-bottom:3px;">{l}</div>
        <div style="font-size:12px;font-weight:600;color:var(--color-text-primary);font-family:var(--font-mono);">
        {'%.2fM'%(t/1e6) if t>=1e6 else '%dK'%(t//1000) if t>=1000 else str(t)}</div>
        {diff_html}</div>"""
    html += '</div>'
    return html

# ── Sidebar ──
with st.sidebar:
    st.markdown("## Inventory Hub")
    week_label = st.text_input("Semana", value="WK13", key="week_input")
    st.markdown(f"**Semana: {week_label}**")
    st.markdown("---")
    st.markdown("**Sube todos los archivos aquí**")
    st.caption("CartonReport → HN | Carton_Report_TLP → TLP | Open_order → Open Order | data_honduras → HN anterior | data_TLP → TLP anterior")
    all_files = st.file_uploader("Selecciona todos los archivos CSV", type=['csv'],
                                  accept_multiple_files=True, key='all_files')
    # Auto-detect files by name
    carton_file = open_file = prev_hn_file = tlp_file = prev_tlp_file = None
    for f in (all_files or []):
        n = f.name.lower()
        if 'carton' in n and 'tlp' in n:      tlp_file = f
        elif 'carton' in n:                    carton_file = f
        elif 'open' in n and 'order' in n:     open_file = f
        elif 'open_order' in n:                open_file = f
        elif ('data_hn' in n or 'data_honduras' in n or 'inventario_hn' in n):  prev_hn_file = f
        elif ('data_tlp' in n or 'inventario_tlp' in n):                         prev_tlp_file = f
    # Show detected files
    if all_files:
        st.markdown("**Archivos detectados:**")
        def check(f, label):
            if f: st.markdown(f"<div style='font-size:11px;color:#4CAF50;'>✓ {label}: {f.name}</div>", unsafe_allow_html=True)
            else: st.markdown(f"<div style='font-size:11px;color:#9CA3AF;'>— {label}: no detectado</div>", unsafe_allow_html=True)
        check(carton_file, "Carton HN")
        check(open_file,   "Open Order")
        check(tlp_file,    "Carton TLP")
        check(prev_hn_file,"HN anterior")
        check(prev_tlp_file,"TLP anterior")
    st.markdown("---")
    if st.button("Clasificar ambas bodegas", type="primary", use_container_width=True):
        if carton_file and open_file:
            with st.spinner("Clasificando Honduras..."):
                carton_file.seek(0); open_file.seek(0)
                r_hn, cut = classify_honduras(pd.read_csv(carton_file,low_memory=False), pd.read_csv(open_file,low_memory=False))
                st.session_state['hn_r'] = r_hn; st.session_state['hn_cut'] = cut
        if prev_hn_file:
            prev_hn_file.seek(0)
            raw = pd.read_csv(prev_hn_file, low_memory=False, header=None)
            if is_pivot_format(raw):
                parsed = parse_prev_hn(raw)
                st.session_state['hn_prev_df'] = None
                st.session_state['hn_prev_clas'] = parsed or {}
            else:
                prev_hn_file.seek(0)
                st.session_state['hn_prev_df'] = pd.read_csv(prev_hn_file, low_memory=False)
                st.session_state['hn_prev_clas'] = None
        else:
            st.session_state['hn_prev_df'] = None
            st.session_state['hn_prev_clas'] = None
        if tlp_file:
            with st.spinner("Clasificando TLP..."):
                tlp_file.seek(0); st.session_state['tlp_r'] = classify_tlp(pd.read_csv(tlp_file,low_memory=False))
        if prev_tlp_file:
            prev_tlp_file.seek(0)
            raw2 = pd.read_csv(prev_tlp_file, low_memory=False, header=None)
            if is_pivot_format(raw2):
                parsed2 = parse_prev_tlp(raw2)
                st.session_state['tlp_prev_df'] = None
                st.session_state['tlp_prev_clas'] = parsed2 or {}
            else:
                prev_tlp_file.seek(0)
                st.session_state['tlp_prev_df'] = pd.read_csv(prev_tlp_file, low_memory=False)
                st.session_state['tlp_prev_clas'] = None
        else:
            st.session_state['tlp_prev_df'] = None
            st.session_state['tlp_prev_clas'] = None
        st.success("Listo!")

# ── Tabs ──
tab_dash, tab_hn, tab_tlp, tab_comp, tab_dl = st.tabs(["Dashboard","Honduras","TLP","Comparativo","Descargas"])

r_hn   = st.session_state.get('hn_r')
r_tlp  = st.session_state.get('tlp_r')
cut_n, cut_u = st.session_state.get('hn_cut',(0,0))
prev_hn  = st.session_state.get('hn_prev_df')
prev_tlp = st.session_state.get('tlp_prev_df')

# ══ TAB DASHBOARD ══
with tab_dash:
    if r_hn is None and r_tlp is None:
        st.info("Carga los archivos en el panel izquierdo y presiona Clasificar.")
    else:
        col_hn, col_tlp = st.columns(2)
        with col_hn:
            if r_hn is not None:
                st.markdown("#### Honduras")
                view_hn = st.radio("", ["Todo","Finished Goods","Wip"], horizontal=True, key="dash_hn_view", label_visibility="collapsed")
                vmap = {"Todo":"all","Finished Goods":"fg","Wip":"wip"}
                df_v = filter_df(r_hn, vmap[view_hn], HN_FG_CLAS, HN_WIP_CLAS)
                cs = df_v.groupby('Clasificacion')['Quantity'].sum().sort_values(ascending=False)
                st.plotly_chart(make_donut(cs, CLAS_COLORS_HN, ""), use_container_width=True)
                st.markdown(color_detail(cs, CLAS_COLORS_HN), unsafe_allow_html=True)
            else:
                st.info("Carga los archivos de Honduras.")
        with col_tlp:
            if r_tlp is not None:
                st.markdown("#### TLP")
                view_tlp = st.radio("", ["Todo","Finished Goods","Wip"], horizontal=True, key="dash_tlp_view", label_visibility="collapsed")
                df_v2 = filter_df(r_tlp, vmap[view_tlp], TLP_FG_CLAS, TLP_WIP_CLAS)
                cs2 = df_v2.groupby('Clasificacion')['Quantity'].sum().sort_values(ascending=False)
                st.plotly_chart(make_donut(cs2, CLAS_COLORS_TLP, ""), use_container_width=True)
                st.markdown(color_detail(cs2, CLAS_COLORS_TLP), unsafe_allow_html=True)
            else:
                st.info("Carga el archivo TLP.")

# ══ TAB HONDURAS ══
with tab_hn:
    if r_hn is None:
        st.info("Carga el Carton Report y Open Order en el panel izquierdo.")
    else:
        st.markdown(f"### Honduras — {week_label}")
        view_h = st.radio("", ["Todo","Finished Goods","Wip"], horizontal=True, key="hn_view", label_visibility="collapsed")
        vmap2 = {"Todo":"all","Finished Goods":"fg","Wip":"wip"}
        df_h = filter_df(r_hn, vmap2[view_h], HN_FG_CLAS, HN_WIP_CLAS)
        df_h['Quantity'] = pd.to_numeric(df_h['Quantity'], errors='coerce').fillna(0)

        tot_uds = int(df_h['Quantity'].sum())
        tot_caj = len(df_h)
        n_clas  = df_h['Clasificacion'].nunique()
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.markdown(kpi_card("Total Cajas", fmt(tot_caj), view_h), unsafe_allow_html=True)
        with c2: st.markdown(kpi_card("Total Unidades", fmt(tot_uds), view_h), unsafe_allow_html=True)
        with c3: st.markdown(kpi_card("Clasificaciones", str(n_clas), "tipos activos",'#185FA5','#378ADD'), unsafe_allow_html=True)
        with c4: st.markdown(kpi_card("Cut excluidas", fmt(cut_n), f"{fmt(cut_u)} uds",'#A32D2D','#E24B4A'), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Por cliente — mayor a menor")
        all_clas_h = HN_FG_CLAS if vmap2[view_h]=='fg' else HN_WIP_CLAS if vmap2[view_h]=='wip' else HN_FG_CLAS+HN_WIP_CLAS
        render_client_table(df_h, all_clas_h, CLAS_COLORS_HN, '#1B5E20', '#F0FDF4')

# ══ TAB TLP ══
with tab_tlp:
    if r_tlp is None:
        st.info("Carga el Carton Report TLP en el panel izquierdo.")
    else:
        st.markdown(f"### TLP — {week_label}")
        view_t = st.radio("", ["Todo","Finished Goods","Wip"], horizontal=True, key="tlp_view", label_visibility="collapsed")
        df_t = filter_df(r_tlp, vmap2[view_t], TLP_FG_CLAS, TLP_WIP_CLAS)
        df_t['Quantity'] = pd.to_numeric(df_t['Quantity'], errors='coerce').fillna(0)

        tot_uds2 = int(df_t['Quantity'].sum())
        tot_caj2 = len(df_t)
        c1,c2,c3 = st.columns(3)
        with c1: st.markdown(kpi_card("Total Cajas", fmt(tot_caj2), view_t,'#0C447C','#378ADD'), unsafe_allow_html=True)
        with c2: st.markdown(kpi_card("Total Unidades", fmt(tot_uds2), view_t,'#0C447C','#378ADD'), unsafe_allow_html=True)
        with c3: st.markdown(kpi_card("Clientes", str(df_t['Customer Name'].nunique()), "con inventario",'#185FA5','#378ADD'), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Por cliente — mayor a menor")
        all_clas_t = TLP_FG_CLAS if vmap2[view_t]=='fg' else TLP_WIP_CLAS if vmap2[view_t]=='wip' else TLP_FG_CLAS+TLP_WIP_CLAS
        render_client_table(df_t, all_clas_t, CLAS_COLORS_TLP, '#0C447C', '#EBF3FB')

# ══ TAB COMPARATIVO ══
with tab_comp:
    st.markdown("### Comparativo histórico")
    st.caption("Sube archivos de semanas anteriores para acumular el historial. Soporta WK1 a WK52.")

    # Multi-file uploader for historical weeks
    hist_files = st.file_uploader(
        "Sube archivos de semanas anteriores (CSV data completa o resumen pivot)",
        type=['csv'], accept_multiple_files=True, key='hist_files'
    )

    # Process historical files and add to session state
    if hist_files:
        for hf in hist_files:
            hname = hf.name.lower()
            # Detect week label from filename e.g. WK12, wk12, semana12
            import re
            wk_match = re.search(r'wk(\d+)', hname) or re.search(r'semana(\d+)', hname) or re.search(r'(\d+)', hname)
            wk_lbl = f"WK{wk_match.group(1)}" if wk_match else hf.name[:10]

            hf.seek(0)
            raw = pd.read_csv(hf, low_memory=False, header=None)

            # Detect bodega and format
            is_hn  = 'tlp' not in hname
            is_tlp = 'tlp' in hname

            if is_pivot_format(raw):
                if is_hn:
                    parsed = parse_prev_hn(raw)
                    if parsed:
                        hist_key = 'hist_hn'
                        if hist_key not in st.session_state: st.session_state[hist_key] = {}
                        st.session_state[hist_key][wk_lbl] = parsed
                else:
                    parsed = parse_prev_tlp(raw)
                    if parsed:
                        hist_key = 'hist_tlp'
                        if hist_key not in st.session_state: st.session_state[hist_key] = {}
                        st.session_state[hist_key][wk_lbl] = parsed
            else:
                hf.seek(0)
                df_hist = pd.read_csv(hf, low_memory=False)
                if 'Clasificacion' in df_hist.columns and 'Quantity' in df_hist.columns:
                    df_hist['Quantity'] = pd.to_numeric(df_hist['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    clas_totals = df_hist.groupby('Clasificacion')['Quantity'].sum().to_dict()
                    hist_key = 'hist_tlp' if is_tlp else 'hist_hn'
                    if hist_key not in st.session_state: st.session_state[hist_key] = {}
                    st.session_state[hist_key][wk_lbl] = clas_totals

    comp_tab_hn, comp_tab_tlp = st.tabs(["Honduras","TLP"])

    def render_comp(bodega_key, r_cur, fg_clas, wip_clas, clas_colors, line_color, week_lbl):
        view_c = st.radio("Ver:", ["Total","Finished Goods","Wip"], horizontal=True,
                          key=f"comp_{bodega_key}_view", label_visibility="visible")
        vmap3 = {"Total":"all","Finished Goods":"fg","Wip":"wip"}

        # Session storage for historical weeks
        hist_key = f'hist_{bodega_key}'
        if hist_key not in st.session_state:
            st.session_state[hist_key] = {}

        # Add current week from classified data
        if r_cur is not None:
            df_c = filter_df(r_cur, vmap3[view_c], fg_clas, wip_clas).copy()
            df_c['Quantity'] = pd.to_numeric(df_c['Quantity'], errors='coerce').fillna(0)
            cur_clas = df_c.groupby('Clasificacion')['Quantity'].sum().to_dict()
            st.session_state[hist_key][week_lbl] = cur_clas

        # Add prev week from parsed pivot if available
        prev_clas_direct = st.session_state.get(f'{bodega_key}_prev_clas')
        if prev_clas_direct:
            prev_wk = f"WK{int(week_lbl.replace('WK',''))-1}" if 'WK' in week_lbl and week_lbl[2:].isdigit() else "WK Ant."
            if prev_wk not in st.session_state[hist_key]:
                st.session_state[hist_key][prev_wk] = prev_clas_direct

        hist = st.session_state[hist_key]

        # Sort weeks properly WK1..WK52
        def wk_sort(w):
            m = re.search(r'(\d+)', w)
            return int(m.group(1)) if m else 999
        weeks_labels = sorted(hist.keys(), key=wk_sort)
        weeks_data   = [hist[w] for w in weeks_labels]
        totals = [sum(d.values()) for d in weeks_data]

        if not weeks_labels:
            st.info("Clasifica el inventario y sube archivos de semanas anteriores para ver el historial.")
            return

        # Badges
        st.markdown(wk_badges_html(totals, weeks_labels, line_color), unsafe_allow_html=True)

        if len(weeks_labels) > 1:
            # KPIs
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            diff = totals[-1]-totals[0]
            dc = '#1B5E20' if diff>=0 else '#A32D2D'
            with kpi_col1: st.markdown(kpi_card("Semana actual", f"{totals[-1]:,}", week_lbl, line_color, line_color), unsafe_allow_html=True)
            with kpi_col2: st.markdown(kpi_card("Primera semana", f"{totals[0]:,}", weeks_labels[0], line_color, line_color), unsafe_allow_html=True)
            with kpi_col3: st.markdown(kpi_card("Variación total", f"{'+' if diff>=0 else ''}{diff:,}", f"{weeks_labels[0]} → {weeks_labels[-1]}", dc, dc), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # Line chart
            fig, _ = render_line_chart(weeks_data, weeks_labels, line_color, bodega_key)
            st.plotly_chart(fig, use_container_width=True)

        # Table
        st.markdown("#### Desglose por clasificación")
        all_clas_c = sorted(set(k for d in weeks_data for k in d.keys()))
        rows = []
        for c in all_clas_c:
            row = {'Clasificación': c}
            for i,(w,d) in enumerate(zip(weeks_labels,weeks_data)):
                val = int(d.get(c,0))
                prev_val = int(weeks_data[i-1].get(c,0)) if i>0 else None
                if prev_val is not None and val != prev_val:
                    arr = '▲' if val > prev_val else '▼'
                    row[w] = f"{val:,} {arr}"
                else:
                    row[w] = f"{val:,}" if val else "—"
            if len(weeks_data) > 1:
                diff_c = int(weeks_data[-1].get(c,0)) - int(weeks_data[0].get(c,0))
                row['Var. total'] = f"{'+' if diff_c>=0 else ''}{diff_c:,}"
            rows.append(row)

        if rows:
            tot_row = {'Clasificación': 'Total'}
            for i,(w,d) in enumerate(zip(weeks_labels,weeks_data)):
                tot_row[w] = f"{int(sum(d.values())):,}"
            if len(weeks_data) > 1:
                tot_diff = int(sum(weeks_data[-1].values())) - int(sum(weeks_data[0].values()))
                tot_row['Var. total'] = f"{'+' if tot_diff>=0 else ''}{tot_diff:,}"
            rows.append(tot_row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # Alerts
        if len(weeks_labels) > 1:
            st.markdown("---")
            st.markdown("#### Alertas de atención")
            prev_df_alert = st.session_state.get(f'{bodega_key}_prev_df')
            r_alert = r_cur
            if r_alert is not None and prev_df_alert is not None:
                build_alerts(r_alert, prev_df_alert, clas_colors)
            elif len(weeks_data) > 1:
                st.info("Sube el CSV de la semana anterior para ver las alertas por cliente.")

        # Button to clear history
        if st.button(f"Limpiar historial {bodega_key.upper()}", key=f"clear_{bodega_key}"):
            st.session_state[hist_key] = {}
            st.rerun()

    with comp_tab_hn:
        if r_hn is None:
            st.info("Clasifica Honduras primero.")
        else:
            render_comp('hn', r_hn, HN_FG_CLAS, HN_WIP_CLAS, CLAS_COLORS_HN, '#1B5E20', week_label)

    with comp_tab_tlp:
        if r_tlp is None:
            st.info("Clasifica TLP primero.")
        else:
            render_comp('tlp', r_tlp, TLP_FG_CLAS, TLP_WIP_CLAS, CLAS_COLORS_TLP, '#0C447C', week_label)


# ══ TAB DESCARGAS ══
with tab_dl:
    st.markdown("### Descargar archivos")
    if r_hn is None and r_tlp is None:
        st.info("Clasifica primero los inventarios para poder descargar.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Honduras")
            if r_hn is not None:
                wk_prev_hn = None
                if prev_hn is not None and 'Clasificacion' in prev_hn.columns and 'Customer Name' in prev_hn.columns:
                    prev_hn4 = prev_hn.copy()
                    prev_hn4['Quantity'] = pd.to_numeric(prev_hn4['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    inv_types = ['Regulars','VMI','Excess','Irregulars','Obsolete','Liability']
                    prev_hn4['Clas_Col'] = prev_hn4['Clasificacion'].map(CMAP_HN)
                    pfg = prev_hn4[prev_hn4['Type']=='Finished Goods'] if 'Type' in prev_hn4.columns else prev_hn4
                    pp = pfg.pivot_table(index='Customer Name',columns='Clas_Col',values='Quantity',aggfunc='sum',fill_value=0)
                    wk_prev_hn = {cli:{t:int(pp.loc[cli,t]) if cli in pp.index and t in pp.columns else 0 for t in inv_types} for cli in ACTIVOS+INACTIVOS}
                with st.spinner("Generando Excel..."):
                    buf_hn = build_excel_hn(r_hn, wk_prev_hn, week_label)
                st.download_button(f"Excel Honduras {week_label}", data=buf_hn,
                    file_name=f"inventario_honduras_{week_label}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                b2 = io.StringIO(); r_hn.to_csv(b2, index=False)
                st.download_button("CSV Honduras completo", data=b2.getvalue(),
                    file_name=f"data_honduras_{week_label}.csv", mime="text/csv",
                    use_container_width=True)
            else:
                st.info("Clasifica Honduras primero.")

        with c2:
            st.markdown("#### TLP")
            if r_tlp is not None:
                wk_prev_tlp = None
                if prev_tlp is not None and 'Clasificacion' in prev_tlp.columns and 'Customer Name' in prev_tlp.columns:
                    prev_tlp4 = prev_tlp.copy()
                    prev_tlp4['Quantity'] = pd.to_numeric(prev_tlp4['Quantity'].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                    inv_comp = ['TLP Irregulars','TLP Printed Excess','TLP sin clasificacion','TLP Blanks Excess']
                    pp2 = prev_tlp4.pivot_table(index='Customer Name',columns='Clasificacion',values='Quantity',aggfunc='sum',fill_value=0)
                    all_tlp_c = list(set(TLP_ORDER+list(pp2.index)))
                    wk_prev_tlp = {cli:{t:int(pp2.loc[cli,t]) if cli in pp2.index and t in pp2.columns else 0 for t in inv_comp} for cli in all_tlp_c}
                with st.spinner("Generando Excel..."):
                    buf_tlp = build_excel_tlp(r_tlp, wk_prev_tlp, week_label)
                st.download_button(f"Excel TLP {week_label}", data=buf_tlp,
                    file_name=f"inventario_TLP_{week_label}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                b4 = io.StringIO(); r_tlp.to_csv(b4, index=False)
                st.download_button("CSV TLP completo", data=b4.getvalue(),
                    file_name=f"data_TLP_{week_label}.csv", mime="text/csv",
                    use_container_width=True)
            else:
                st.info("Clasifica TLP primero.")
