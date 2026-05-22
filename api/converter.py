import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io
from pathlib import Path
from http.server import BaseHTTPRequestHandler
import json
import base64

def limpar(texto):
    return re.sub(r'\s*P[aá]gina:.*$', '', texto).strip()

def processar(conteudo_bytes, nome_arquivo):
    registros = []
    cnpj = razao = comp = 'N/A'
    cc = 'Sem CC'
    tipo = None
    ignorar = False

    with pdfplumber.open(io.BytesIO(conteudo_bytes)) as pdf:
        for pagina in pdf.pages:
            for linha in (pagina.extract_text() or "").split('\n'):
                linha = linha.strip()
                if not linha: continue
                m = re.match(r'Empresa:\s*\d+\s*-\s*(.+)', linha, re.I)
                if m: razao = limpar(m.group(1)); continue
                m = re.match(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', linha, re.I)
                if m: cnpj = m.group(1); continue
                m = re.match(r'Compet.ncia:\s*(\d{2}/\d{4})', linha, re.I)
                if m: comp = m.group(1); continue
                m = re.match(r'C\.?Custo:\s*(.+)', linha, re.I)
                if m: cc = m.group(1).strip(); tipo = None; ignorar = False; continue
                if re.search(r'Resumo das bases', linha, re.I): tipo = None; ignorar = True; continue
                if ignorar: continue
                if re.match(r'^(Folha Mensal|Total:|Liquido|Rubrica)\s*', linha, re.I): continue
                if re.match(r'^PROVENTOS\s*$', linha, re.I): tipo = 'Provento'; continue
                if re.match(r'^DESCONTOS\s*$', linha, re.I): tipo = 'Desconto'; continue
                if re.match(r'^INFORMATIVA\s*$', linha, re.I): tipo = 'Informativo'; continue
                m = re.match(r'^(\d{1,6})\s+(.+?)\s+[\d:]+\s+[\d:.,]+\s+([\d.]+,\d{2})\*?\s*$', linha)
                if m and tipo:
                    v = float(m.group(3).replace('.','').replace(',','.'))
                    registros.append({
                        'CNPJ': cnpj, 'Razao Social': razao, 'Competencia': comp,
                        'Centro de Custo': cc, 'Tipo': tipo,
                        'Rubrica': m.group(1)+' - '+m.group(2).strip(),
                        'Valor': v, 'Arquivo': nome_arquivo
                    })
    return registros

CAB="1F4E79"; PRO="C6EFCE"; DES="FFC7CE"; INF="FFEB9C"; TOT="D9E1F2"; CC_COR="DEEAF1"

def ecab(c):
    c.font=Font(bold=True,color="FFFFFF",size=11)
    c.fill=PatternFill("solid",start_color=CAB)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    b=Side(style="thin",color="AAAAAA"); c.border=Border(left=b,right=b,top=b,bottom=b)

def edat(c,cor=None):
    b=Side(style="thin",color="DDDDDD"); c.border=Border(left=b,right=b,top=b,bottom=b)
    c.alignment=Alignment(vertical="center")
    if cor: c.fill=PatternFill("solid",start_color=cor)

def ctipo(t):
    return {'Provento':PRO,'Desconto':DES,'Informativo':INF}.get(t)

def aba_cnpj(wb, cnpj, df):
    df = df.groupby(['CNPJ','Razao Social','Competencia','Centro de Custo','Tipo','Rubrica'], sort=False).agg(Valor=('Valor','sum')).reset_index()
    ws = wb.create_sheet(title=re.sub(r'[^\w]','',cnpj)[:31])
    razao = df['Razao Social'].iloc[0]; comp = df['Competencia'].iloc[0]
    ws.merge_cells('A1:D1'); ws['A1']='FOLHA DE PAGAMENTO - '+razao
    ws['A1'].font=Font(bold=True,size=13,color="1F4E79"); ws['A1'].alignment=Alignment(horizontal="center")
    ws.merge_cells('A2:D2'); ws['A2']='CNPJ: '+cnpj+'  |  Competencia: '+comp
    ws['A2'].font=Font(italic=True,color="444444"); ws['A2'].alignment=Alignment(horizontal="center")
    hdrs=['Centro de Custo','Tipo','Rubrica','Valor (R$)']; wds=[35,14,50,16]
    for col,(h,w) in enumerate(zip(hdrs,wds),1):
        c=ws.cell(row=4,column=col,value=h); ecab(c)
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[4].height=22
    linha=5
    for cc, dcc in df.groupby('Centro de Custo',sort=False):
        ws.merge_cells('A'+str(linha)+':D'+str(linha))
        c=ws.cell(row=linha,column=1,value='  '+cc)
        c.font=Font(bold=True,color="1F4E79",size=10)
        c.fill=PatternFill("solid",start_color=CC_COR)
        c.alignment=Alignment(horizontal="left",vertical="center")
        ws.row_dimensions[linha].height=18; linha+=1
        tot={'Provento':0.0,'Desconto':0.0,'Informativo':0.0}
        for _,row in dcc.iterrows():
            cor=ctipo(row['Tipo'])
            for col,val in enumerate([cc,row['Tipo'],row['Rubrica'],row['Valor']],1):
                c=ws.cell(row=linha,column=col,value=val); edat(c,cor)
                if col==4: c.number_format='#,##0.00'; c.alignment=Alignment(horizontal="right")
            tot[row['Tipo']]=tot.get(row['Tipo'],0)+row['Valor']; linha+=1
        res='Proventos: R$ {:,.2f}  |  Descontos: R$ {:,.2f}  |  Informativos: R$ {:,.2f}'.format(tot['Provento'],tot['Desconto'],tot['Informativo'])
        for col in range(1,5):
            c=ws.cell(row=linha,column=col,value=res if col==1 else '')
            c.fill=PatternFill("solid",start_color=TOT); c.font=Font(bold=True)
            if col==1: c.alignment=Alignment(horizontal="left",indent=1)
        ws.row_dimensions[linha].height=16; linha+=2
    ws.freeze_panes='A5'

def aba_resumo(wb, df):
    ws=wb.create_sheet(title="RESUMO",index=0)
    ws.merge_cells('A1:F1'); ws['A1']="RESUMO CONSOLIDADO - FOLHA DE PAGAMENTO"
    ws['A1'].font=Font(bold=True,size=14,color="1F4E79"); ws['A1'].alignment=Alignment(horizontal="center")
    hdrs=['CNPJ','Razao Social','Competencia','Total Proventos','Total Descontos','Total Informativos']; wds=[22,42,14,18,18,18]
    for col,(h,w) in enumerate(zip(hdrs,wds),1):
        c=ws.cell(row=3,column=col,value=h); ecab(c)
        ws.column_dimensions[get_column_letter(col)].width=w
    linha=4
    for (cnpj,comp),grp in df.groupby(['CNPJ','Competencia']):
        razao=grp['Razao Social'].iloc[0]
        for col,val in enumerate([cnpj,razao,comp,grp[grp.Tipo=='Provento'].Valor.sum(),grp[grp.Tipo=='Desconto'].Valor.sum(),grp[grp.Tipo=='Informativo'].Valor.sum()],1):
            c=ws.cell(row=linha,column=col,value=val); edat(c)
            if col>=4: c.number_format='R$ #,##0.00'; c.alignment=Alignment(horizontal="right")
        linha+=1
    for col in range(1,7):
        c=ws.cell(row=linha,column=col); c.fill=PatternFill("solid",start_color=TOT); c.font=Font(bold=True)
    ws.merge_cells('A'+str(linha)+':C'+str(linha))
    ws.cell(row=linha,column=1,value="TOTAL GERAL").alignment=Alignment(horizontal="center")
    for col,tipo in zip([4,5,6],['Provento','Desconto','Informativo']):
        c=ws.cell(row=linha,column=col,value=df[df.Tipo==tipo].Valor.sum())
        c.number_format='R$ #,##0.00'; c.font=Font(bold=True); c.alignment=Alignment(horizontal="right")
    ws.freeze_panes='A4'

def gerar_excel(registros):
    df = pd.DataFrame(registros)
    wb = Workbook(); wb.remove(wb.active)
    aba_resumo(wb, df)
    for cnpj, dfc in df.groupby('CNPJ'):
        aba_cnpj(wb, cnpj, dfc)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/api/converter':
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            try:
                todos = []
                for arq in body['arquivos']:
                    dados = base64.b64decode(arq['conteudo'])
                    todos.extend(processar(dados, arq['nome']))
                excel = gerar_excel(todos)
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename=folha_pagamento.xlsx')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(excel)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'erro': str(e)}).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
