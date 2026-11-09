from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

wb=Workbook()
F='Arial'
H=Font(name=F,bold=True,color='FFFFFF'); HF=PatternFill('solid',fgColor='245B4A')
B=Font(name=F,bold=True); N=Font(name=F); BLUE=Font(name=F,color='0000FF'); GREEN=Font(name=F,color='008000')
Y=PatternFill('solid',fgColor='FFFF00'); FIX=PatternFill('solid',fgColor='FCE4D6')
thin=Side(style='thin',color='BBBBBB'); BD=Border(left=thin,right=thin,top=thin,bottom=thin)
NUM='#,##0;(#,##0);-'; PCT='0.0%'
def hdr(ws,r,vals,c0=1):
    for i,v in enumerate(vals):
        c=ws.cell(r,c0+i,v); c.font=H; c.fill=HF; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BD
def cell(ws,r,c,v,font=N,fmt=None,fill=None,note=None):
    x=ws.cell(r,c,v); x.font=font; x.border=BD
    if fmt: x.number_format=fmt
    if fill: x.fill=fill
    if note: x.comment=Comment(note,'Claude')
    return x
def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=Font(name=F,bold=True,size=14)
    if sub: ws['A2']=sub; ws['A2'].font=Font(name=F,italic=True,color='666666')

# ---------- Sheet: README
rd=wb.active; rd.title='อ่านก่อน'
title(rd,'CAPEX CMNH — ฉบับแก้ไข (รพ.เฉพาะทางระบบประสาท สมอง และไขสันหลัง เชียงใหม่ มาสเตอร์พีซ)','สร้างใหม่จากไฟล์ CAPEX CMNH.xlsx (Drive) วันที่ 9 พ.ย. 2569')
lines=[('สีของเซลล์',''),('ตัวอักษรสีน้ำเงิน','ตัวเลขที่ป้อนเอง (แก้ได้)'),('ตัวอักษรสีดำ','สูตร คำนวณอัตโนมัติ'),('ตัวอักษรสีเขียว','ดึงค่าจากชีตอื่น'),('พื้นเหลือง','ยังไม่มีข้อมูล ต้องกรอก (เงินลงทุน รายได้ OPD ฯลฯ)'),('พื้นส้มอ่อน','จุดที่แก้จากไฟล์เดิม (มีคอมเมนต์บอกค่าเดิม)'),('',''),('สิ่งที่แก้',''),
('1','ราคา Villa T.Tube ราคา 3: 120,000 → 101,000 (ตามระยะห่างราคา 2 → 3 ของกลุ่มอื่น ยืนยันอีกครั้ง)'),
('2','ตารางเวรโซนรวม: หัวคอลัมน์ Day/Day แก้เป็น Day/Night และช่วงเตียง Villa 1–4, 4–8 แก้เป็น 1–4, 5–8'),
('3','ผู้ช่วยพยาบาล IPD 32 → คำนวณจากเวรต่อวัน × 30/26 (มีคนสำรองวันหยุด) = 37 คน; Villa 10 → 12 คน'),
('4','OPD เพิ่มผู้ช่วยเหลือคนไข้ (NA) 2 คน ที่อยู่ในตารางแผนกแต่ขาดในตารางค่าจ้าง'),
('5','IPD เพิ่มนักกายภาพ 1 คน คุม PTA สำหรับแพ็กเกจ Rehab'),
('6','ยอดรวมทุกตารางเปลี่ยนเป็นสูตร และเพิ่มชีต "รายได้และจุดคุ้มทุน" ที่มีช่องกรอกเงินลงทุนและเงินกู้'),
('7','อัตราค่าจ้างใช้ค่าสูงสุดของช่วงตามไฟล์เดิม (เช่น ผู้ช่วยพยาบาล 15,000)'),('8','9 พ.ย. 69: ปรับบุคลากรตามคำขอใบอนุญาต (แพทย์ 4, RN 15, เภสัชกร 3, นักเทคนิคการแพทย์ 3, รังสีเทคนิค 2, OT 1) และใส่เงินลงทุน 161.8 ล้านบาท เงินกู้ 66% ดอกเบี้ย 6% 15 ปี (ดอกเบี้ยและระยะเวลาเป็นสมมติฐาน)')]
for i,(a,b) in enumerate(lines):
    rd.cell(4+i,1,a).font=B if b=='' else N; rd.cell(4+i,2,b).font=N
rd['A5'].font=BLUE; rd['A6'].font=N; rd['A7'].font=GREEN; rd['A8'].fill=Y; rd['A9'].fill=FIX
rd.column_dimensions['A'].width=22; rd.column_dimensions['B'].width=95

# ---------- Sheet: prices
ps=wb.create_sheet('ราคาแพ็กเกจ'); title(ps,'ราคาแพ็กเกจรายเดือน (บาท/เตียง/เดือน)')
hdr(ps,4,['ประเภทห้อง','จำนวนเตียง','กลุ่มอาการ','ราคาขาย 1','ราคาขาย 2','ราคาขาย 3','หมายเหตุ'])
rooms=[('โซนรวม',38,[(45000,47000,49000),(45000,47000,49000),(49000,52000,54000),(58000,59000,62000),(62500,65000,68000)]),
('ห้องพิเศษคู่',10,[(52000,55000,58000),(52000,55000,58000),(59000,65000,67000),(68000,70000,75000),(72000,78000,82000)]),
('ห้องพิเศษเดี่ยว',2,[(65000,67000,69000),(65000,67000,69000),(72000,75000,78000),(78000,80000,82000),(82500,85000,87000)]),
('Villa',20,[(75000,79000,82000),(75000,79000,82000),(85000,89000,92000),(92500,95000,97000),(97000,99000,101000)])]
groups=['ผู้สูงอายุ','Alzheimer','Rehabilitation','Feeding','T.Tube']
notes={'ผู้สูงอายุ':'','Alzheimer':'ไม่รวมจนท.ประกบ','Rehabilitation':'รวมกายภาพ PTA 6/week','Feeding':'รวมเปลี่ยนสาย NG','T.Tube':'รวมเปลี่ยนสาย NG / care TT Tube ไม่รวมจนท.ประกบ'}
r=5; base_rows={}
for name,beds,pr in rooms:
    for g,p in zip(groups,pr):
        cell(ps,r,1,name); cell(ps,r,2,beds,BLUE); cell(ps,r,3,g)
        for j,v in enumerate(p):
            c=cell(ps,r,4+j,v,BLUE,NUM)
        n=notes[g] if not (name=='โซนรวม' and g=='Alzheimer') else ''
        cell(ps,r,7,n)
        if g=='ผู้สูงอายุ': base_rows[name]=r
        if name=='Villa' and g=='T.Tube':
            ps.cell(r,6).fill=FIX; ps.cell(r,6).comment=Comment('ไฟล์เดิม 120,000 กระโดดจาก 99,000 น่าจะพิมพ์ผิด ปรับเป็น 101,000 ตามระยะห่างราคาของกลุ่มอื่น กรุณายืนยัน','Claude')
        r+=1
r+=1; ps.cell(r,1,'แพ็กเกจเสริม (บวกเพิ่มจากค่าห้อง)').font=B; r+=1
hdr(ps,r,['แพ็กเกจ','ราคา (บาท)','รายละเอียด']); r+=1
addons=[('กายภาพหลังผ่าตัดกระดูก',25000,'คอร์ส 30 วัน: Hydro 10 ครั้ง PT 10 ครั้ง Ultrasound 10 ครั้ง'),
('ยากระตุ้นสมอง (ยาฉีด)',35000,'ยาฉีด 1 คอร์ส 15 วัน'),('ยากระตุ้นสมอง (ยากิน)',17000,'1 คอร์ส'),
('Post Stroke 1',69000,'เคสไม่ตื่น ไม่สามารถทำตามสั่ง'),('Post Stroke 2',89000,'เคสเริ่มตื่น ทำตามสั่งได้บนเตียง'),
('Post Stroke 3',113000,'เคสตื่น ทำตามสั่งได้'),('Post Stroke 4',125000,'เคสตื่น ทำตามสั่งได้ ทรงตัวได้'),
('TMS 10 ครั้ง',22000,''),('TMS + PMS + ฝังเข็ม 10 ครั้ง',42000,''),('กายภาพบำบัด 30 ครั้ง',24000,''),('ธาราบำบัด 10 ครั้ง',12000,'')]
for a,p,d in addons:
    cell(ps,r,1,a); cell(ps,r,2,p,BLUE,NUM); cell(ps,r,3,d); r+=1
for c,w in zip('ABCDEFG',[24,12,16,13,13,13,44]): ps.column_dimensions[c].width=w
ps.freeze_panes='A5'

# ---------- Sheet: staffing
st=wb.create_sheet('อัตรากำลังเวร'); title(st,'ผู้ช่วยพยาบาล (PN/NA) ต่อเวร ตามจำนวนผู้พัก','จำนวนคนที่ต้องจ้าง = เวรต่อวันตอนเต็ม × ตัวคูณสำรองวันหยุด')
st['A4']='ตัวคูณสำรองวันหยุด (30 วัน / 26 วันทำงาน)'; st['A4'].font=B
cell(st,4,5,'=30/26',N,'0.000',note='สมมติพนักงานทำงาน 26 วัน/เดือน ปรับได้')
zones=[('โซนรวมชาย 19 เตียง',['1–5','6–10','11–15','16–19'],[(2,2),(3,2),(4,3),(5,4)]),
('โซนรวมหญิง 19 เตียง',['1–5','6–10','11–15','16–19'],[(2,2),(3,2),(4,3),(5,4)]),
('ห้องพิเศษคู่ 5 ห้อง 10 เตียง',['1–2','3–4','5–6','7–8','9–10'],[(1,1),(2,2),(3,3),(4,4),(5,5)]),
('ห้องพิเศษเดี่ยว 2 ห้อง 2 เตียง',['1','2'],[(1,1),(2,2)])]
r=6; ipd_full=[]
for name,ranges,vals in zones:
    st.cell(r,1,name).font=B; r+=1
    hdr(st,r,['ผู้พัก (เตียง)']+sum([[g,''] for g in ranges],[]))
    for i in range(len(ranges)): st.merge_cells(start_row=r,start_column=2+2*i,end_row=r,end_column=3+2*i)
    r+=1; hdr(st,r,['']+['Day','Night']*len(ranges))
    if name.startswith('โซนรวม'):
        for cc in (6,7): st.cell(r,cc).fill=FIX
        st.cell(r,7).comment=Comment('ไฟล์เดิมหัวคอลัมน์เขียน Day/Day แก้เป็น Day/Night','Claude')
    r+=1; cell(st,r,1,'ผู้ช่วยพยาบาล (PN NA)')
    for i,(d,n) in enumerate(vals):
        cell(st,r,2+2*i,d,BLUE); cell(st,r,3+2*i,n,BLUE)
    lastc=1+2*len(vals)
    cell(st,r,lastc+1,f'=SUM({L(lastc-1)}{r}:{L(lastc)}{r})',B); st.cell(r-1,lastc+1,'เวร/วัน ตอนเต็ม').font=B
    ipd_full.append(f'{L(lastc+1)}{r}'); r+=2
st.cell(r,1,'รวม IPD: เวร/วัน ตอนเต็ม').font=B; cell(st,r,5,'='+'+'.join(ipd_full),B); ipd_shift=r; r+=1
st.cell(r,1,'ผู้ช่วยพยาบาล IPD ที่ต้องจ้าง (คน)').font=B; cell(st,r,5,f'=ROUNDUP(E{ipd_shift}*$E$4,0)',B,fill=FIX,note='ไฟล์เดิมจ้าง 32 คน = เวรต่อวันพอดี ไม่มีสำรองวันหยุด'); ipd_need=r; r+=2
st.cell(r,1,'Villa 5 หลัง 20 ห้อง (ไม่รวมการดูแลแบบ 1:1)').font=B; r+=1
vr=['1–4','5–8','9–12','13–16','17–20']
hdr(st,r,['ผู้พัก (เตียง)']+sum([[g,''] for g in vr],[]))
for i in range(5): st.merge_cells(start_row=r,start_column=2+2*i,end_row=r,end_column=3+2*i)
st.cell(r,4).fill=FIX; st.cell(r,4).comment=Comment('ไฟล์เดิมเขียน 4-8 ซ้อนกับ 1-4 แก้เป็น 5-8','Claude')
r+=1; hdr(st,r,['']+['Day','Night']*5); r+=1
cell(st,r,1,'ผู้ช่วยพยาบาล (PN NA)')
for i in range(5): cell(st,r,2+2*i,i+1,BLUE); cell(st,r,3+2*i,i+1,BLUE)
cell(st,r,12,f'=SUM(J{r}:K{r})',B); st.cell(r-1,12,'เวร/วัน ตอนเต็ม').font=B; vshift=r; r+=1
st.cell(r,1,'ผู้ช่วยพยาบาล Villa ที่ต้องจ้าง (คน)').font=B; cell(st,r,5,f'=ROUNDUP(L{vshift}*$E$4,0)',B,fill=FIX,note='ไฟล์เดิมจ้าง 10 คน'); v_need=r
st.column_dimensions['A'].width=40
for c in range(2,13): st.column_dimensions[L(c)].width=9

# ---------- Sheet: wages
wg=wb.create_sheet('ค่าจ้าง'); title(wg,'อัตราจ้างงานรายเดือน (บาท)')
roles=['แพทย์','พยาบาลวิชาชีพ','ผู้ดำเนินการ (care manager)','ผู้ช่วยพยาบาล','ผู้ช่วยเหลือคนไข้ (NA)','นักกายภาพ','ผู้ช่วยนักกายภาพ','แม่บ้าน','แม่ครัว','การเงิน การบัญชี','เวรเปล','เภสัชกร','นักเทคนิคการแพทย์','นักรังสีเทคนิค','นักกิจกรรมบำบัด','ยาม','ฝ่ายการตลาด + Content + Admin']
rng={'แพทย์':'ขั้นต้น 135,000','พยาบาลวิชาชีพ':'22,000–25,000','ผู้ดำเนินการ (care manager)':'22,000–25,000','ผู้ช่วยพยาบาล':'12,000–15,000','ผู้ช่วยเหลือคนไข้ (NA)':'12,000–15,000','นักกายภาพ':'20,000–25,000','ผู้ช่วยนักกายภาพ':'12,000–15,000','แม่บ้าน':'10,000–12,000','แม่ครัว':'10,000–12,000','การเงิน การบัญชี':'15,000–20,000','เวรเปล':'10,000–12,000','เภสัชกร':'35,000','นักเทคนิคการแพทย์':'25,000 (สมมติ)','นักรังสีเทคนิค':'25,000 (สมมติ)','นักกิจกรรมบำบัด':'25,000 (สมมติ)','ยาม':'10,000–12,000','ฝ่ายการตลาด + Content + Admin':'25,000 + commission (Villa 12,000)'}
rate={'แพทย์':135000,'พยาบาลวิชาชีพ':25000,'ผู้ดำเนินการ (care manager)':25000,'ผู้ช่วยพยาบาล':15000,'ผู้ช่วยเหลือคนไข้ (NA)':15000,'นักกายภาพ':25000,'ผู้ช่วยนักกายภาพ':15000,'แม่บ้าน':12000,'แม่ครัว':12000,'การเงิน การบัญชี':20000,'เวรเปล':12000,'เภสัชกร':35000,'นักเทคนิคการแพทย์':25000,'นักรังสีเทคนิค':25000,'นักกิจกรรมบำบัด':25000,'ยาม':12000,'ฝ่ายการตลาด + Content + Admin':25000}
hc={'OPD':{'แพทย์':2,'พยาบาลวิชาชีพ':5,'ผู้ช่วยพยาบาล':7,'ผู้ช่วยเหลือคนไข้ (NA)':2,'นักกายภาพ':3,'ผู้ช่วยนักกายภาพ':2,'แม่บ้าน':2,'การเงิน การบัญชี':1,'เวรเปล':3,'เภสัชกร':2,'ยาม':1,'ฝ่ายการตลาด + Content + Admin':1},
'IPD':{'แพทย์':2,'พยาบาลวิชาชีพ':10,'ผู้ช่วยพยาบาล':'IPDNEED','นักกายภาพ':1,'ผู้ช่วยนักกายภาพ':3,'แม่บ้าน':3,'แม่ครัว':2,'การเงิน การบัญชี':2,'เภสัชกร':1,'นักเทคนิคการแพทย์':3,'นักรังสีเทคนิค':2,'นักกิจกรรมบำบัด':1,'ยาม':1},
'Villa':{'ผู้ดำเนินการ (care manager)':1,'ผู้ช่วยพยาบาล':'VNEED','แม่บ้าน':2,'แม่ครัว':1,'การเงิน การบัญชี':1,'ยาม':1,'ฝ่ายการตลาด + Content + Admin':1}}
fixnotes={('OPD','ผู้ช่วยเหลือคนไข้ (NA)'):'เพิ่มใหม่: ตารางแผนก OPD มี NA 2 คน (ห้องตรวจ 1, ห้องผ่าตัด) แต่ไฟล์เดิมไม่มีในตารางค่าจ้าง',
('IPD','นักกายภาพ'):'เพิ่มใหม่ 1 คน: IPD ขายแพ็กเกจ Rehab (PTA 6 ครั้ง/สัปดาห์) ต้องมีนักกายภาพคุม PTA ไฟล์เดิมเป็น 0',
('IPD','แพทย์'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: แพทย์ 4 = OPD 2 + IPD 2) ไฟล์เดิม IPD 1',('IPD','พยาบาลวิชาชีพ'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: RN 15 = OPD 5 + IPD 10) ไฟล์เดิม IPD 4',('IPD','เภสัชกร'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: เภสัชกร 3 = OPD 2 + IPD 1) ไฟล์เดิม IPD 0',('IPD','นักเทคนิคการแพทย์'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: นักเทคนิคการแพทย์ 3) ไฟล์เดิมไม่มี',('IPD','นักรังสีเทคนิค'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: รังสีเทคนิค 2 สำหรับ CT) ไฟล์เดิมไม่มี',('IPD','นักกิจกรรมบำบัด'):'ปรับตามคำขอใบอนุญาต ส.พ.4 / ตารางบุคลากร ส.พ.6 (ทั้ง รพ.: OT 1) ไฟล์เดิมไม่มี',('IPD','ผู้ช่วยพยาบาล'):'ไฟล์เดิม 32 คน คำนวณใหม่ในชีตอัตรากำลังเวร',('Villa','ผู้ช่วยพยาบาล'):'ไฟล์เดิม 10 คน คำนวณใหม่ในชีตอัตรากำลังเวร'}
r=4; totals={}; heads={}
for z in ['OPD','IPD','Villa']:
    lab={'OPD':'OPD คลินิกเวชกรรม (08.00–20.00)','IPD':'IPD 50 เตียง','Villa':'Nursing Home Villa 20 ห้อง'}[z]
    wg.cell(r,1,lab).font=Font(name=F,bold=True,size=12); r+=1
    hdr(wg,r,['ตำแหน่ง','ช่วงเงินเดือน','จำนวน (คน)','ค่าจ้าง/คน','รวม']); r+=1; first=r
    for role in roles:
        if role not in hc[z]: continue
        v=hc[z][role]
        cell(wg,r,1,role); cell(wg,r,2,rng[role])
        if v=='IPDNEED': c=cell(wg,r,3,f"='อัตรากำลังเวร'!E{ipd_need}",GREEN)
        elif v=='VNEED': c=cell(wg,r,3,f"='อัตรากำลังเวร'!E{v_need}",GREEN)
        else: c=cell(wg,r,3,v,BLUE)
        rt=12000 if (z=='Villa' and role.startswith('ฝ่ายการตลาด')) else rate[role]
        cell(wg,r,4,rt,BLUE,NUM); cell(wg,r,5,f'=C{r}*D{r}',N,NUM)
        if (z,role) in fixnotes: c.fill=FIX; c.comment=Comment(fixnotes[(z,role)],'Claude')
        r+=1
    cell(wg,r,1,'รวม',B); cell(wg,r,3,f'=SUM(C{first}:C{r-1})',B); cell(wg,r,5,f'=SUM(E{first}:E{r-1})',B,NUM)
    totals[z]=f'E{r}'; heads[z]=f'C{r}'; r+=2
wg.cell(r,1,'รวมทุกโซน').font=B; cell(wg,r,3,'='+'+'.join(heads.values()),B); cell(wg,r,5,'='+'+'.join(totals.values()),B,NUM); wtot=r
for c,w in zip('ABCDE',[34,30,13,13,14]): wg.column_dimensions[c].width=w

# ---------- Sheet: total cost
tc=wb.create_sheet('ค่าใช้จ่ายรวม'); title(tc,'ค่าใช้จ่ายดำเนินงานรายเดือน (บาท)','ยังไม่รวมเงินลงทุนและค่าผ่อนธนาคาร (ดูชีตรายได้และจุดคุ้มทุน)')
hdr(tc,4,['รายการ','บาท/เดือน','หมายเหตุ'])
cell(tc,5,1,'ค่าจ้าง OPD'); cell(tc,5,2,f"='ค่าจ้าง'!{totals['OPD']}",GREEN,NUM)
cell(tc,6,1,'ค่าจ้าง IPD'); cell(tc,6,2,f"='ค่าจ้าง'!{totals['IPD']}",GREEN,NUM)
cell(tc,7,1,'ค่าจ้าง Villa'); cell(tc,7,2,f"='ค่าจ้าง'!{totals['Villa']}",GREEN,NUM)
other=[('ค่าเช่าที่',120000,''),('ค่าไฟ',150000,''),('ค่าน้ำ',50000,''),('ค่าซักล้าง (70 คน)',200000,''),('อาหาร (70 คน)',450000,''),('เช่าบริการ เช่น กำจัดแมลง น้ำหอม กำจัดยุง',30000,''),('ค่าซ่อมแซม maintenance',30000,''),('ค่าเวชภัณฑ์ (รวมในค่าบริการ)',50000,'')]
r=8
for a,v,n in other: cell(tc,r,1,a); cell(tc,r,2,v,BLUE,NUM); cell(tc,r,3,n); r+=1
cell(tc,r,1,'รวมค่าใช้จ่ายรายเดือน',B); cell(tc,r,2,f'=SUM(B5:B{r-1})',B,NUM); mt=r; r+=1
cell(tc,r,1,'รวมค่าใช้จ่ายรายปี',B); cell(tc,r,2,f'=B{mt}*12',B,NUM); r+=1
cell(tc,r,1,'จำนวนผู้พักเต็ม (เตียง)'); cell(tc,r,2,70,BLUE,NUM); bedsr=r; r+=1
cell(tc,r,1,'ต้นทุนต่อผู้พัก/เดือน (ที่เต็ม 70 เตียง)',B); cell(tc,r,2,f'=B{mt}/B{bedsr}',B,NUM); r+=2
tc.cell(r,1,'เทียบไฟล์เดิม: ค่าใช้จ่ายรายเดือน 2,999,000 บาท').font=Font(name=F,italic=True,color='666666')
for c,w in zip('ABC',[44,16,30]): tc.column_dimensions[c].width=w

# ---------- Sheet: revenue & break-even
rv=wb.create_sheet('รายได้และจุดคุ้มทุน'); title(rv,'รายได้ค่าเตียง จุดคุ้มทุน และภาระเงินกู้','รายได้ค่าเตียงคิดที่ราคากลุ่มผู้สูงอายุ (ต่ำสุด) — กลุ่มอาการอื่นแพงกว่า')
rv['A4']='เลือกระดับราคา (1, 2 หรือ 3)'; rv['A4'].font=B; cell(rv,4,2,1,BLUE)
hdr(rv,6,['โซน','เตียง','ราคา/เตียง','รายได้เต็ม 100%'])
zr=[('โซนรวม','โซนรวม'),('ห้องพิเศษคู่','ห้องพิเศษคู่'),('ห้องพิเศษเดี่ยว','ห้องพิเศษเดี่ยว'),('Villa','Villa')]
r=7
for z,_ in zr:
    br=base_rows[z]
    cell(rv,r,1,z); cell(rv,r,2,f"='ราคาแพ็กเกจ'!B{br}",GREEN)
    cell(rv,r,3,f"=INDEX('ราคาแพ็กเกจ'!D{br}:F{br},1,$B$4)",GREEN,NUM)
    cell(rv,r,4,f'=B{r}*C{r}',N,NUM); r+=1
cell(rv,r,1,'รวม',B); cell(rv,r,2,f'=SUM(B7:B{r-1})',B); cell(rv,r,4,f'=SUM(D7:D{r-1})',B,NUM); fullrev=r; r+=2
rv.cell(r,1,'ช่องที่ต้องกรอก').font=B; r+=1
inputs=[('รายได้ OPD ต่อเดือน (บาท)','ไฟล์เดิมไม่มี กรอกจากประมาณการผู้ป่วยนอก'),('รายได้แพ็กเกจเสริมต่อเดือน (บาท)','Post Stroke, TMS+PMS, ยากระตุ้นสมอง ฯลฯ'),('เงินลงทุนรวม CAPEX (บาท)','ค่าก่อสร้าง DG + อุปกรณ์ + ตกแต่ง'),('วงเงินกู้ธนาคาร (บาท)',''),('ดอกเบี้ยต่อปี',''),('ระยะเวลากู้ (ปี)','')]
inp={}
DEF_IN={'เงินลงทุนรวม CAPEX (บาท)':('=132306570+12500000+17000000','ส.พ.4: ก่อสร้าง 132,306,570 (รวม VAT) + ตกแต่ง 10–15 ล้าน (ใช้ 12.5) + เครื่องมือแพทย์ 14–20 ล้าน (ใช้ 17)'),'วงเงินกู้ธนาคาร (บาท)':(None,'ส.พ.4: สถาบันการเงิน 66% (แบบฟอร์มรวม 110% ต้องแก้)'),'ดอกเบี้ยต่อปี':(0.06,'สมมติ รอเงื่อนไขธนาคารจริง'),'ระยะเวลากู้ (ปี)':(15,'สมมติ รอเงื่อนไขธนาคารจริง')}
for a,n in inputs:
    cell(rv,r,1,a); dv=DEF_IN.get(a,(None,n))
    if a=='วงเงินกู้ธนาคาร (บาท)': dv=(f"=B{inp['เงินลงทุนรวม CAPEX (บาท)']}*0.66",dv[1])
    c=cell(rv,r,2,dv[0],BLUE,PCT if 'ดอกเบี้ย' in a else NUM,Y); cell(rv,r,3,dv[1] or n); inp[a]=r; r+=1
r+=1
cell(rv,r,1,'ค่าผ่อนเงินกู้ต่อเดือน (บาท)',B)
lr,ir,tr=inp['วงเงินกู้ธนาคาร (บาท)'],inp['ดอกเบี้ยต่อปี'],inp['ระยะเวลากู้ (ปี)']
cell(rv,r,2,f'=IF(OR(B{lr}="",B{ir}="",B{tr}=""),0,-PMT(B{ir}/12,B{tr}*12,B{lr}))',B,NUM); pay=r; r+=1
cell(rv,r,1,'ค่าใช้จ่ายดำเนินงานต่อเดือน (บาท)'); cell(rv,r,2,f"='ค่าใช้จ่ายรวม'!B{mt}",GREEN,NUM); opx=r; r+=2
rv.cell(r,1,'จุดคุ้มทุน (อัตราเข้าพักที่ต้องมี)').font=B; r+=1
op,ad=inp['รายได้ OPD ต่อเดือน (บาท)'],inp['รายได้แพ็กเกจเสริมต่อเดือน (บาท)']
cell(rv,r,1,'ก่อนค่าผ่อนเงินกู้'); cell(rv,r,2,f'=IF(D{fullrev}=0,0,(B{opx}-N(B{op})-N(B{ad}))/D{fullrev})',B,PCT); r+=1
cell(rv,r,1,'หลังค่าผ่อนเงินกู้'); cell(rv,r,2,f'=IF(D{fullrev}=0,0,(B{opx}+B{pay}-N(B{op})-N(B{ad}))/D{fullrev})',B,PCT); r+=2
rv.cell(r,1,'กำไร/ขาดทุนต่อเดือน ตามอัตราเข้าพัก').font=B; r+=1
hdr(rv,r,['อัตราเข้าพัก','รายได้ค่าเตียง','รายได้ OPD + เสริม','ค่าใช้จ่ายดำเนินงาน','ค่าผ่อนเงินกู้','กำไร/ขาดทุน']); r+=1
for occ in [0.6,0.7,0.8,0.9,1.0]:
    cell(rv,r,1,occ,BLUE,PCT); cell(rv,r,2,f'=$D${fullrev}*A{r}',N,NUM); cell(rv,r,3,f'=N($B${op})+N($B${ad})',N,NUM)
    cell(rv,r,4,f'=$B${opx}',N,NUM); cell(rv,r,5,f'=$B${pay}',N,NUM); cell(rv,r,6,f'=B{r}+C{r}-D{r}-E{r}',B,NUM); r+=1
r+=1
rv.cell(r,1,'หมายเหตุ: ค่าใช้จ่ายในตารางนี้คงที่ตอนเต็ม 70 เตียง ช่วงเปิดใหม่ที่ผู้พักน้อย จ้างคนตามตารางเวร จะต่ำกว่านี้').font=Font(name=F,italic=True,color='666666')
for c,w in zip('ABCDEF',[38,16,20,20,16,16]): rv.column_dimensions[c].width=w
wb.save('CAPEX_CMNH_ฉบับแก้ไข.xlsx')
