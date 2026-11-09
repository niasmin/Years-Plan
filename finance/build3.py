exec(open('base.py').read())
from openpyxl.chart import LineChart, Reference
# locate rows in wage sheet
zone=None; pnrow={}
for row in wg.iter_rows(min_row=1,max_row=wg.max_row):
    v=row[0].value
    if isinstance(v,str) and v.startswith('OPD'): zone='OPD'
    elif isinstance(v,str) and v.startswith('IPD'): zone='IPD'
    elif isinstance(v,str) and v.startswith('Nursing'): zone='Villa'
    if v=='ผู้ช่วยพยาบาล' and zone: pnrow[zone]=row[0].row
TC="'ค่าใช้จ่ายรวม'"; WG="'ค่าจ้าง'"; RV="'รายได้และจุดคุ้มทุน'"
MON=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']

# ================= Monthly plan
mp=wb.create_sheet('แผนรายเดือน 2570-71',1)
title(mp,'แผนการเงินรายเดือน ปี 2570–2571 (บาท)','ช่องสีน้ำเงิน/พื้นเหลือง = สมมติฐาน แก้ได้ ตัวเลขทั้งหมดคำนวณจากชีตอื่นในไฟล์นี้')
inp_rows=[
('เดือนเปิดคลินิกที่อาคารหน้า รพ. (1 = ม.ค. 2570)',1,None,'คลินิกแพทย์ + กายภาพบำบัด เปิดก่อน แล้วย้ายเข้า รพ.'),
('เดือนเปิด OPD คลินิกเวชกรรม (1 = ม.ค. 2570)',10,None,'เท่ากับเดือนที่ย้ายคลินิกเข้า รพ. (พร้อมเปิด IPD)'),
('เดือนเปิด IPD 50 เตียง',10,None,'สมมติ: วางโครงสร้าง รพ. Q3 แล้วเปิด ต.ค. — ต้องยืนยัน'),
('เดือนเปิด Villa 20 ห้อง',11,None,'แผน Q4: เริ่ม operate Villa'),
('จ้างพนักงานล่วงหน้าก่อนเปิด (เดือน)',1,None,'จ่ายเงินเดือนก่อนเปิดเพื่อฝึกงาน'),
('รายได้คลินิกหน้า รพ.ต่อเดือน',0,NUM,'ยังไม่มีประมาณการ — ดูช่อง "รายได้คลินิกหน้า รพ.ขั้นต่ำ" ด้านล่าง'),
('ค่าจ้างคลินิกหน้า รพ./เดือน',272000,NUM,'แพทย์ 1 · RN 1 · PN 1 · PT 2 · PTA 1 · ต้อนรับ/การเงิน 1 · แม่บ้าน 1 = 8 คน (อัตราเดียวกับชีตค่าจ้าง) · ทีมนี้ย้ายเข้า OPD รพ.'),
('ค่าใช้จ่ายอื่นคลินิกหน้า รพ./เดือน',30000,NUM,'สมมติ: น้ำไฟ วัสดุการแพทย์ ขยะติดเชื้อ (สถานที่เป็นของบริษัท ไม่มีค่าเช่า)'),
('ค่าตกแต่งอาคารหน้า รพ. เป็นคลินิก (ครั้งเดียว)',200000,NUM,'สมมติ — ให้ผู้รับเหมาประเมินจริง · ลงในเดือนแรก'),
('รายได้ OPD ต่อเดือน หลังเปิด',0,NUM,'ยังไม่มีประมาณการ — ดูช่อง "รายได้ OPD ขั้นต่ำ" ด้านล่าง'),
('รายได้แพ็กเกจเสริมต่อผู้พัก/เดือน (เฉลี่ย)',0,NUM,'Post Stroke, TMS+PMS ฯลฯ ใส่ค่าเฉลี่ยต่อหัว'),
('ผู้ช่วยพยาบาล IPD ขั้นต่ำเมื่อเปิด (คน)',21,None,'ตามใบอนุญาต: PN 20 + NA 10 = 30 คน หักที่ OPD มีแล้ว 9 คน'),
('ผู้ช่วยพยาบาล Villa ขั้นต่ำเมื่อเปิด (คน)',3,None,''),
('สัดส่วนค่าน้ำไฟที่เกิดตั้งแต่เปิด OPD',0.5,PCT,'ที่เหลือเกิดเมื่อเปิด IPD'),
('เปิด OPD แบบเล็กก่อนเปิด IPD (1 = ใช่ · 0 = เปิดเต็มตั้งแต่แรก)',0,None,'ทางเลือก B = 1'),
('ค่าจ้าง OPD แบบเล็ก/เดือน',386000,NUM,'1 ห้องตรวจ 08.00–17.00: แพทย์ 1 · RN 2 · PN 3 · เภสัชกร 1 · PT 1 · PTA 1 · การเงิน 1 · แม่บ้าน 1 · ยาม 1 · Admin/การตลาด 1 · เวรเปล 1'),
('เดือนเริ่มผ่อนธนาคาร',10,None,'สมมติเริ่มผ่อนเมื่อเปิด IPD (ต.ค. 70) — ยืนยันกับธนาคาร'),
('ผู้พักออก/เสียชีวิต ต่อเดือน (% ของผู้พัก)',0.04,PCT,'สมมติฐาน ใช้คำนวณจำนวนรับใหม่'),
('อัตราปิดการขาย (ผู้สนใจ → เข้าพัก)',0.25,PCT,'สมมติฐาน'),
]
r=4; IN={}
for a,v,fmt,n in inp_rows:
    cell(mp,r,1,a); c=cell(mp,r,2,v,BLUE,fmt,Y); cell(mp,r,3,n); IN[a]=r; r+=1
def ref(k):
    m=[v for kk,v in IN.items() if kk.startswith(k)]
    assert len(m)==1,(k,m)
    return f'$B${m[0]}'
r+=1
mp.cell(r,1,'อัตราเข้าพักตามเดือนที่เปิด (IPD และ Villa)').font=B; ramp_r=r+1
ramp=[0.15,0.25,0.35,0.45,0.55,0.62,0.68,0.73,0.77,0.80,0.82,0.85]
cell(mp,ramp_r,1,'เดือนที่เปิด →',B)
for i,v in enumerate(ramp):
    cell(mp,ramp_r-0,3+i,v,BLUE,PCT,Y); mp.cell(ramp_r-1,3+i,f'เดือน {i+1}').font=B
cell(mp,ramp_r,2,'',N); mp.cell(ramp_r+1,1,'หลังเดือนที่ 12 ใช้ค่าเดือนที่ 12').font=Font(name=F,italic=True,color='666666')
RAMP=f'$C${ramp_r}:$N${ramp_r}'
r=ramp_r+3
# derived constants
mp.cell(r,1,'ค่าที่ดึงมาจากชีตอื่น').font=B; r+=1
der=[('ราคาเฉลี่ย IPD ต่อเตียง/เดือน',f"=SUMPRODUCT({RV}!B7:B9,{RV}!C7:C9)/SUM({RV}!B7:B9)"),
('ราคา Villa ต่อเตียง/เดือน',f"={RV}!C10"),
('ค่าจ้าง OPD/เดือน',f"={WG}!{totals['OPD']}"),
('ค่าจ้าง IPD ส่วนคงที่ (ไม่รวมผู้ช่วยพยาบาล)',f"={WG}!{totals['IPD']}-{WG}!E{pnrow['IPD']}"),
('ค่าจ้าง Villa ส่วนคงที่',f"={WG}!{totals['Villa']}-{WG}!E{pnrow['Villa']}"),
('เงินเดือนผู้ช่วยพยาบาล/คน',f"={WG}!D{pnrow['IPD']}"),
('ผู้ช่วยพยาบาล IPD ตอนเต็ม (คน)',f"={WG}!C{pnrow['IPD']}"),
('ผู้ช่วยพยาบาล Villa ตอนเต็ม (คน)',f"={WG}!C{pnrow['Villa']}"),
('ค่าสถานที่คงที่ (เช่า กำจัดแมลง ซ่อม)',f"={TC}!B8+{TC}!B13+{TC}!B14"),
('ค่าน้ำไฟเต็ม',f"={TC}!B9+{TC}!B10"),
('ค่าแปรผันต่อผู้พัก (อาหาร ซักล้าง เวชภัณฑ์)',f"=({TC}!B11+{TC}!B12+{TC}!B15)/70"),
('ค่าผ่อนธนาคาร/เดือน',f"={RV}!B{pay}")]
D={}
for a,fml in der:
    cell(mp,r,1,a); cell(mp,r,2,fml,GREEN,NUM); D[a]=f'$B${r}'; r+=1
def d(k):
    m=[v for kk,v in D.items() if kk.startswith(k)]
    assert len(m)==1,(k,m)
    return m[0]
r+=1
# grid
g0=r; hdr(mp,g0,['รายการ','หน่วย']+[f'{MON[i%12]} {70+i//12}' for i in range(24)])
mp.row_dimensions[g0].height=30
rows={}
def grow(key,label,unit,fn,fmt=NUM,bold=False):
    global r
    r+=1; rows[key]=r
    cell(mp,r,1,label,B if bold else N); cell(mp,r,2,unit)
    for i in range(24):
        col=L(3+i); prev=L(2+i)
        cell(mp,r,3+i,fn(col,i,prev),B if bold else N,fmt)
r=g0
grow('idx','ลำดับเดือน','',lambda c,i,p:i+1,'0')
R=lambda k,c: f'{c}{rows[k]}'
grow('mi','IPD เปิดมาแล้ว (เดือน)','เดือน',lambda c,i,p:f"=MAX(0,{R('idx',c)}-{ref('เดือนเปิด IPD')}+1)",'0')
grow('mv','Villa เปิดมาแล้ว (เดือน)','เดือน',lambda c,i,p:f"=MAX(0,{R('idx',c)}-{ref('เดือนเปิด Villa')}+1)",'0')
grow('oi','อัตราเข้าพัก IPD','%',lambda c,i,p:f"=IF({R('mi',c)}=0,0,INDEX({RAMP},1,MIN(12,{R('mi',c)})))",PCT)
grow('ov','อัตราเข้าพัก Villa','%',lambda c,i,p:f"=IF({R('mv',c)}=0,0,INDEX({RAMP},1,MIN(12,{R('mv',c)})))",PCT)
grow('ri','ผู้พัก IPD','คน',lambda c,i,p:f"=ROUND(50*{R('oi',c)},0)",'0')
grow('rvl','ผู้พัก Villa','คน',lambda c,i,p:f"=ROUND(20*{R('ov',c)},0)",'0')
grow('rt','ผู้พักรวม','คน',lambda c,i,p:f"={R('ri',c)}+{R('rvl',c)}",'0',True)
r+=1; mp.cell(r,1,'รายได้').font=B
grow('revi','รายได้ IPD','บาท',lambda c,i,p:f"={R('ri',c)}*{d('ราคาเฉลี่ย IPD')}")
grow('revv','รายได้ Villa','บาท',lambda c,i,p:f"={R('rvl',c)}*{d('ราคา Villa')}")
grow('reva','รายได้แพ็กเกจเสริม','บาท',lambda c,i,p:f"={R('rt',c)}*{ref('รายได้แพ็กเกจเสริม')}")
grow('revc','รายได้คลินิกหน้า รพ.','บาท',lambda c,i,p:f"=IF(AND({R('idx',c)}>={ref('เดือนเปิดคลินิกที่อาคารหน้า')},{R('idx',c)}<{ref('เดือนเปิด OPD')}),{ref('รายได้คลินิกหน้า รพ.')},0)")
grow('revo','รายได้ OPD','บาท',lambda c,i,p:f"=IF({R('idx',c)}>={ref('เดือนเปิด OPD')},{ref('รายได้ OPD ต่อ')},0)")
grow('rev','รายได้รวม','บาท',lambda c,i,p:f"=SUM({c}{rows['revi']}:{c}{rows['revo']})",NUM,True)
r+=1; mp.cell(r,1,'กำลังคนและค่าใช้จ่าย').font=B
act=lambda c,k: f"{R('idx',c)}>={ref(k)}-{ref('จ้างพนักงานล่วงหน้า')}"
grow('hi','ผู้ช่วยพยาบาล IPD','คน',lambda c,i,p:f"=IF({act(c,'เดือนเปิด IPD')},MAX({ref('ผู้ช่วยพยาบาล IPD ขั้น')},ROUNDUP({d('ผู้ช่วยพยาบาล IPD ตอน')}*{R('oi',c)},0)),0)",'0')
grow('hv','ผู้ช่วยพยาบาล Villa','คน',lambda c,i,p:f"=IF({act(c,'เดือนเปิด Villa')},MAX({ref('ผู้ช่วยพยาบาล Villa ข')},ROUNDUP({d('ผู้ช่วยพยาบาล Villa ตอ')}*{R('ov',c)},0)),0)",'0')
grow('co','ค่าจ้าง OPD','บาท',lambda c,i,p:f"=IF({act(c,'เดือนเปิด OPD')},IF(AND({ref('เปิด OPD แบบเล็ก')}=1,{R('idx',c)}<{ref('เดือนเปิด IPD')}-{ref('จ้างพนักงานล่วงหน้า')}),{ref('ค่าจ้าง OPD แบบเล็ก')},{d('ค่าจ้าง OPD')}),0)")
grow('cc','ค่าใช้จ่ายคลินิกหน้า รพ.','บาท',lambda c,i,p:f"=IF(AND({R('idx',c)}>={ref('เดือนเปิดคลินิกที่อาคารหน้า')},{R('idx',c)}<{ref('เดือนเปิด OPD')}),{ref('ค่าจ้างคลินิกหน้า รพ.')}+{ref('ค่าใช้จ่ายอื่นคลินิกหน้า รพ.')},0)+IF({R('idx',c)}=MAX(1,{ref('เดือนเปิดคลินิกที่อาคารหน้า')}),{ref('ค่าตกแต่งอาคารหน้า')},0)")
grow('ci','ค่าจ้าง IPD','บาท',lambda c,i,p:f"=IF({act(c,'เดือนเปิด IPD')},{d('ค่าจ้าง IPD')},0)+{R('hi',c)}*{d('เงินเดือนผู้ช่วย')}")
grow('cv','ค่าจ้าง Villa','บาท',lambda c,i,p:f"=IF({act(c,'เดือนเปิด Villa')},{d('ค่าจ้าง Villa')},0)+{R('hv',c)}*{d('เงินเดือนผู้ช่วย')}")
grow('cf','ค่าสถานที่และน้ำไฟ','บาท',lambda c,i,p:f"=IF({R('idx',c)}>={ref('เดือนเปิด OPD')},{d('ค่าสถานที่คงที่')}+{d('ค่าน้ำไฟเต็ม')}*{ref('สัดส่วนค่าน้ำไฟ')},0)+IF({R('mi',c)}>0,{d('ค่าน้ำไฟเต็ม')}*(1-{ref('สัดส่วนค่าน้ำไฟ')}),0)")
grow('cx','ค่าแปรผันตามผู้พัก','บาท',lambda c,i,p:f"={R('rt',c)}*{d('ค่าแปรผันต่อผู้พัก')}")
grow('cl','ค่าผ่อนธนาคาร','บาท',lambda c,i,p:f"=IF({R('idx',c)}>={ref('เดือนเริ่มผ่อน')},{d('ค่าผ่อนธนาคาร')},0)")
grow('cost','ค่าใช้จ่ายรวม','บาท',lambda c,i,p:f"=SUM({c}{rows['co']}:{c}{rows['cl']})",NUM,True)
r+=1
grow('net','กำไร/ขาดทุนรายเดือน','บาท',lambda c,i,p:f"={R('rev',c)}-{R('cost',c)}",NUM,True)
grow('cum','เงินสดสะสม (ต้องเตรียมเงินเมื่อติดลบ)','บาท',lambda c,i,p:(f"={R('net',c)}" if i==0 else f"={p}{rows['cum']}+{R('net',c)}"),NUM,True)
grow('be','(ช่วยคำนวณ) เดือนที่กำไรหลังเปิด IPD','',lambda c,i,p:f"=IF(AND({R('mi',c)}>0,{R('net',c)}>0),{R('idx',c)},999)",'0')
r+=1; mp.cell(r,1,'แผนการขาย').font=B
grow('adm','ผู้พักรับใหม่ที่ต้องได้','คน',lambda c,i,p:(f"={R('rt',c)}" if i==0 else f"=MAX(0,{R('rt',c)}-{p}{rows['rt']})+ROUND({p}{rows['rt']}*{ref('ผู้พักออก')},0)"),'0',True)
grow('lead','ผู้สนใจที่ต้องได้ (lead)','ราย',lambda c,i,p:f"=ROUNDUP({R('adm',c)}/{ref('อัตราปิดการขาย')},0)",'0',True)
last=L(26)
r+=2; mp.cell(r,1,'สรุป').font=Font(name=F,bold=True,size=12); r+=1
S={}
for a,fml,fmt in [('เงินที่ต้องเตรียมสูงสุด (เงินสดสะสมต่ำสุด)',f"=-MIN(0,MIN(C{rows['cum']}:{last}{rows['cum']}))",NUM),
('เดือนแรกที่มีกำไร หลังเปิด IPD (ลำดับเดือน)',f"=IF(MIN(C{rows['be']}:{last}{rows['be']})=999,\"ยังไม่คุ้มทุนใน 24 เดือน\",MIN(C{rows['be']}:{last}{rows['be']}))",'0'),
('กำไร/ขาดทุนรวมปี 2570',f"=SUM(C{rows['net']}:N{rows['net']})",NUM),
('กำไร/ขาดทุนรวมปี 2571',f"=SUM(O{rows['net']}:Z{rows['net']})",NUM),
('รายได้ OPD ขั้นต่ำ/เดือน ให้คุ้มค่าจ้าง OPD + ค่าสถานที่',f"={d('ค่าจ้าง OPD')}+{d('ค่าสถานที่คงที่')}+{d('ค่าน้ำไฟเต็ม')}*{ref('สัดส่วนค่าน้ำไฟ')}",NUM),
('รายได้คลินิกหน้า รพ.ขั้นต่ำ/เดือน ให้คุ้มทุนคลินิก',f"={ref('ค่าจ้างคลินิกหน้า รพ.')}+{ref('ค่าใช้จ่ายอื่นคลินิกหน้า รพ.')}",NUM),
('ค่าใช้จ่ายคลินิกหน้า รพ.รวม ก่อนย้ายเข้า รพ.',f"=SUM(C{rows['cc']}:Z{rows['cc']})",NUM),
('ผู้สนใจที่ต้องได้ทั้งปี 2570',f"=SUM(C{rows['lead']}:N{rows['lead']})",'0'),
('ผู้สนใจที่ต้องได้ทั้งปี 2571',f"=SUM(O{rows['lead']}:Z{rows['lead']})",'0')]:
    cell(mp,r,1,a,B); cell(mp,r,2,fml,B,fmt); S[a]=r; r+=1
r+=1; mp.cell(r,1,'แหล่งที่มาของผู้พักใหม่ (ข้อเสนอ แก้ได้)').font=B; r+=1
hdr(mp,r,['แหล่ง','สัดส่วน','lead ปี 2570','lead ปี 2571']); r+=1; src0=r
for a,v in [('ส่งต่อจากโรงพยาบาลพันธมิตร',0.35),('คลินิกหมอต้น / OPD คลินิกเวชกรรม',0.20),('ย้ายจากสาขา 1–4 และ TQL (ขึ้นห้องพิเศษ/Villa)',0.15),('ออนไลน์และการตลาด',0.20),('ตลาดต่างประเทศ (เริ่ม Q4)',0.10)]:
    cell(mp,r,1,a); cell(mp,r,2,v,BLUE,PCT,Y)
    cell(mp,r,3,f"=ROUND(B{r}*$B${S['ผู้สนใจที่ต้องได้ทั้งปี 2570']},0)",N,'0'); cell(mp,r,4,f"=ROUND(B{r}*$B${S['ผู้สนใจที่ต้องได้ทั้งปี 2571']},0)",N,'0'); r+=1
cell(mp,r,1,'รวม',B); cell(mp,r,2,f'=SUM(B{src0}:B{r-1})',B,PCT)
mp.column_dimensions['A'].width=46; mp.column_dimensions['B'].width=14
for i in range(24): mp.column_dimensions[L(3+i)].width=11.5
mp.freeze_panes=mp.cell(g0+1,3)
ch=LineChart(); ch.title='เงินสดสะสม (บาท)'; ch.height=7; ch.width=22
ch.add_data(Reference(mp,min_col=3,max_col=26,min_row=rows['cum'],max_row=rows['cum']),from_rows=True,titles_from_data=False)
ch.set_categories(Reference(mp,min_col=3,max_col=26,min_row=g0,max_row=g0)); ch.legend=None
mp.add_chart(ch,f'D{S["เงินที่ต้องเตรียมสูงสุด (เงินสดสะสมต่ำสุด)"]}')

# ================= Existing business baseline
eb=wb.create_sheet('ธุรกิจเดิม 2569',2)
title(eb,'ธุรกิจเดิม: ฐานปี 2569 และเป้าปี 2570 (บาท/เดือน)','กรอกพื้นเหลืองจากตัวเลขจริง เพื่อใช้ตั้งเป้า KPI และดูภาพรวมทั้งบริษัท')
hdr(eb,4,['หน่วยธุรกิจ','ความจุ','อัตราเข้าพัก/ใช้งาน 2569','รายได้/เดือน 2569','ค่าใช้จ่าย/เดือน 2569','กำไร/เดือน 2569','เป้าเติบโตรายได้ 2570','กำไรทั้งปี 2570 (ประมาณ)'])
units=[('สาขา 1–4 (Nursing Home)','80 เตียง'),('Tranquila (TQL)','40 เตียง'),('Day Care','10–15 คน'),('Home Care',''),('ให้เช่า/ขายเครื่องมือแพทย์',''),('คลินิกหมอต้น (TMS, PMS, PT)',''),('CM.Derma+','')]
r=5
for a,cap in units:
    cell(eb,r,1,a); cell(eb,r,2,cap)
    for cc,fmt in [(3,PCT),(4,NUM),(5,NUM)]: cell(eb,r,cc,None,BLUE,fmt,Y)
    cell(eb,r,7,0.10,BLUE,PCT,Y,note='เป้าร่าง +10% แก้ได้')
    cell(eb,r,6,f'=IF(D{r}="","",D{r}-E{r})',N,NUM)
    cell(eb,r,8,f'=IF(D{r}="","",(D{r}*(1+N(G{r}))-E{r})*12)',N,NUM,note='สมมติค่าใช้จ่ายคงเดิม รายได้โตตามเป้า')
    r+=1
cell(eb,r,1,'รวมธุรกิจเดิม',B)
for cc in (4,5,6,8): cell(eb,r,cc,f'=SUM({L(cc)}5:{L(cc)}{r-1})',B,NUM)
ebt=r; r+=3
eb.cell(r,1,'ภาพรวมบริษัท ปี 2570').font=Font(name=F,bold=True,size=12); r+=1
hdr(eb,r,['รายการ','บาท']); r+=1
cell(eb,r,1,'กำไรธุรกิจเดิม (ประมาณ)'); cell(eb,r,2,f'=H{ebt}',N,NUM); a1=r; r+=1
cell(eb,r,1,'กำไร/ขาดทุนโครงการ รพ. + Villa'); cell(eb,r,2,f"='แผนรายเดือน 2570-71'!B{S['กำไร/ขาดทุนรวมปี 2570']}",GREEN,NUM); a2=r; r+=1
cell(eb,r,1,'กำไร/ขาดทุนรวมบริษัท',B); cell(eb,r,2,f'=B{a1}+B{a2}',B,NUM); r+=1
cell(eb,r,1,'เงินที่ต้องเตรียมสำหรับโครงการ รพ. (สูงสุด)'); cell(eb,r,2,f"='แผนรายเดือน 2570-71'!B{S['เงินที่ต้องเตรียมสูงสุด (เงินสดสะสมต่ำสุด)']}",GREEN,NUM)
for c,w in zip('ABCDEFGH',[34,12,14,16,16,16,14,18]): eb.column_dimensions[c].width=w

# ================= KPI tracking
kt=wb.create_sheet('ติดตาม KPI 2570',3)
title(kt,'ติดตาม KPI ปี 2570 — เป้าร่างที่กรอกไว้ให้ก่อน (แก้ได้ในช่องเหลือง)','แต่ละ KPI มี 2 แถว: "เป้า" (ร่าง แก้ได้) และ "ผล" (กรอกทุกวันที่ 5 ของเดือนถัดไป) ช่องผลจะเป็นสีเขียวเมื่อถึงเป้า สีแดงเมื่อไม่ถึง · เป้าสีเขียวดึงจากชีตแผนรายเดือนอัตโนมัติ')
hdr(kt,4,['KPI','สาย','หน่วย','ทิศทาง','แถว']+[f'{m} 70' for m in MON]+['ผ่านเป้า (เดือน)','ที่มาของเป้า / หมายเหตุ'])
MP="'แผนรายเดือน 2570-71'!"
def mlink(row,blank0=True):
    return [f'=IF({MP}{L(3+i)}{row}=0,"",{MP}{L(3+i)}{row})' if blank0 else f'={MP}{L(3+i)}{row}' for i in range(12)]
def const(v,start=0):
    return [None]*start+[v]*(12-start)
EB="'ธุรกิจเดิม 2569'!"
rev_t=[f'=IF(SUM({EB}$D$5:$D$11)=0,"",SUMPRODUCT({EB}$D$5:$D$11,1+{EB}$G$5:$G$11)+{MP}{L(3+i)}{rows["rev"]})' for i in range(12)]
clin_t=[f'=IF({MP}{L(3+i)}{rows["cc"]}>0,{MP}$B${S["รายได้คลินิกหน้า รพ.ขั้นต่ำ/เดือน ให้คุ้มทุนคลินิก"]},"")' for i in range(12)]
lead_t=[None]*6+[5,10,20]+[f'={MP}{L(3+i)}{rows["lead"]}' for i in range(9,12)]
opd_t=[40,60,80,100,120,140,160,180,200,300,350,400]
K=[
 ('งานในตารางงานเสร็จตามกำหนด','CEO','%','↑',[None,None,.9,None,None,.9,None,None,.9,None,None,.9],'≥ 90% วัดรายไตรมาส'),
 ('อัตราเข้าพัก สาขา 1–4','COO','%','↑',const(.85),'≥ 85% ปรับตามค่าฐาน 2569'),
 ('อัตราเข้าพัก TQL','COO','%','↑',const(.80),'≥ 80% เพิ่มขึ้นเมื่อมีผู้พักต่างชาติ'),
 ('อัตราเข้าพัก IPD (หลังเปิด)','COO','%','↑',mlink(rows['oi']),'ตามแผนรายเดือน (เปิด ต.ค. 70)'),
 ('อัตราเข้าพัก Villa (หลังเปิด)','COO','%','↑',mlink(rows['ov']),'ตามแผนรายเดือน (เปิด พ.ย. 70)'),
 ('ผู้ป่วยคลินิก (อาคารหน้า รพ. ม.ค.–ก.ย. → OPD รพ. ต.ค.)','COO','ราย/เดือน','↑',opd_t,'สมมติ: คลินิกหน้า รพ. 40 → 200 ราย/เดือน แล้ว OPD รพ. 300 → 400 ให้ CMO/CFO ยืนยัน'),
 ('พลัดตกใหม่','COO','ครั้ง/1,000 วันนอน','↓',const(1.0),'≤ 1.0'),
 ('แผลกดทับใหม่','COO','ครั้ง/1,000 วันนอน','↓',const(1.0),'≤ 1.0'),
 ('อัตราลาออกพนักงาน','COO + HR','%/เดือน','↓',const(.02),'≤ 2% ต่อเดือน (ราว 24% ต่อปี)'),
 ('ผู้ใช้ Day Care / Home Care','COO','ราย','↑',[None]+[None]*11,'ใส่ค่าฐาน ม.ค. แล้วตั้งเป้า +20% ภายใน ธ.ค.'),
 ('ใบอนุญาตวิชาชีพที่หมดอายุขณะทำงาน','COO + HR','คน','↓',const(0),'ต้องเป็น 0 · HR เตือนล่วงหน้า 90 วัน (ดู docs/03)'),
 ('รายได้รวมบริษัท','CFO','บาท','↑',rev_t,'ธุรกิจเดิม × (1 + เป้าเติบโต) จากชีตธุรกิจเดิม + รายได้ รพ./Villa ตามแผน · ว่างจนกว่าจะกรอกชีตธุรกิจเดิม 2569'),
 ('รายได้คลินิกหน้า รพ.','CFO','บาท','↑',clin_t,'≥ ค่าใช้จ่ายคลินิก (จุดคุ้มทุน) ม.ค.–ก.ย. แล้วย้ายเข้า OPD รพ.'),
 ('ค่าใช้จ่ายโครงการ รพ. + Villa','CFO','บาท','↓',mlink(rows['cost']),'ไม่เกินแผนรายเดือน'),
 ('ปิดงบรายเดือนภายในวันที่','CFO','วันที่','↓',const(10),'ภายในวันที่ 10 ของเดือนถัดไป'),
 ('หนี้ค้างชำระเกิน 30 วัน','CFO','% ของรายได้','↓',const(.05),'≤ 5%'),
 ('งบก่อสร้าง รพ. ใช้จริงเทียบแผน','CFO','%','↓',const(1.0),'≤ 100% (เกิน 5% ต้องขออนุมัติ)'),
 ('เงินสดสะสมโครงการ รพ.','CFO','บาท','↑',mlink(rows['cum']),'ไม่ต่ำกว่าแผนรายเดือน'),
 ('ผู้สนใจ (lead) รพ. + Villa','CMO','ราย','↑',lead_t,'ก.ค.–ก.ย. หาก่อนเปิด จากนั้นตามแผนการขาย (96 ราย/ปี)'),
 ('อัตราปิดการขาย','CMO','%','↑',[None]*9+[.25]*3,'≥ 25%'),
 ('คะแนนความพึงพอใจญาติ','CMO','คะแนน (5)','↑',[None,None,4.5,None,None,4.5,None,None,4.5,None,None,4.5],'≥ 4.5 รายไตรมาส'),
 ('เรื่องร้องเรียน','CMO','เรื่อง/เดือน','↓',const(2),'≤ 2'),
 ('ตปท.: ผู้แนะนำ (multipliers) ที่ติดต่อได้ สะสม','CMO','ราย','↑',[None,None,3,None,None,10,None,None,20,None,None,30],'แพทย์ครอบครัว Spitex ศูนย์ความจำ ฯลฯ ผ่าน Swiss Helping Point'),
 ('ตปท.: lead ต่างชาติ','CMO','ราย','↑',[None]*6+[2,3,4,5,5,6],'ทาง A (25 ราย/ปี) · ทาง B ให้ใช้ ต.ค.–ธ.ค. 2–3 ราย/เดือน'),
 ('ตปท.: ผู้พักต่างชาติรับใหม่ สะสม','CMO','คน','↑',[None]*10+[1,3],'ทาง A · ทาง B เริ่มปี 2571 · เป้าระยะยาว 40 คน'),
 ('ตปท.: พนักงาน TQL ผ่านอบรมภาษาอังกฤษพื้นฐาน','COO + HR','%','↑',[None]*5+[.3,None,None,.5,None,None,.7],'30% มิ.ย. · 50% ก.ย. · 70% ธ.ค.'),
]
r=5; LT=PatternFill('solid',fgColor='FFF2CC')
from openpyxl.formatting.rule import FormulaRule
for a,s_,u,dr,tg,src in K:
    fmt=PCT if '%' in u else ('0.0' if 'คะแนน' in u or '1,000' in u else NUM)
    for rr,lab in ((r,'เป้า'),(r+1,'ผล')):
        cell(kt,rr,1,a if lab=='เป้า' else None,B if lab=='เป้า' else N); cell(kt,rr,2,s_ if lab=='เป้า' else None); cell(kt,rr,3,u if lab=='เป้า' else None); cell(kt,rr,4,dr if lab=='เป้า' else None); cell(kt,rr,5,lab)
    for i,v in enumerate(tg):
        if isinstance(v,str) and v.startswith('='): cell(kt,r,6+i,v,GREEN,fmt)
        else: cell(kt,r,6+i,v,BLUE,fmt,LT)
        cell(kt,r+1,6+i,None,BLUE,fmt)
    t=f'F{r}:Q{r}'; a_=f'F{r+1}:Q{r+1}'
    cell(kt,r+1,18,f'=IF(COUNT({a_})=0,"",SUMPRODUCT(({a_}<>"")*ISNUMBER({t})*IF($D${r}="↑",1,0)*({a_}>={t}))+SUMPRODUCT(({a_}<>"")*ISNUMBER({t})*IF($D${r}="↓",1,0)*({a_}<={t}))&" / "&SUMPRODUCT(({a_}<>"")*ISNUMBER({t})))')
    cell(kt,r,19,src)
    up=f'AND(F{r+1}<>"",ISNUMBER(F{r}),IF($D${r}="↑",F{r+1}>=F{r},F{r+1}<=F{r}))'
    dn=f'AND(F{r+1}<>"",ISNUMBER(F{r}),NOT(IF($D${r}="↑",F{r+1}>=F{r},F{r+1}<=F{r})))'
    kt.conditional_formatting.add(f'F{r+1}:Q{r+1}',FormulaRule(formula=[up],fill=PatternFill('solid',fgColor='C6EFCE')))
    kt.conditional_formatting.add(f'F{r+1}:Q{r+1}',FormulaRule(formula=[dn],fill=PatternFill('solid',fgColor='FFC7CE')))
    r+=2
kt.column_dimensions['A'].width=40; kt.column_dimensions['B'].width=10; kt.column_dimensions['C'].width=16; kt.column_dimensions['D'].width=7; kt.column_dimensions['E'].width=6
for i in range(12): kt.column_dimensions[L(6+i)].width=12
kt.column_dimensions['R'].width=12; kt.column_dimensions['S'].width=60
kt.freeze_panes='F5'
# readme additions
rr=rd.max_row+2
rd.cell(rr,1,'ชีตใหม่').font=B
for i,(a,b) in enumerate([('แผนรายเดือน 2570-71','รายได้ ค่าใช้จ่าย กำลังคน เงินสดสะสม แผนการขาย รายเดือน 24 เดือน'),('ธุรกิจเดิม 2569','กรอกตัวเลขธุรกิจเดิม → ภาพรวมทั้งบริษัท'),('ติดตาม KPI 2570','กรอกผลจริงรายเดือน')]):
    rd.cell(rr+1+i,1,a).font=N; rd.cell(rr+1+i,2,b).font=N
rd['A1']='CMNH แผนการเงินปี 2570–71 (รวม CAPEX ฉบับแก้ไข)'
wb.save('CMNH_แผนการเงิน_2570-71.xlsx')
import json; json.dump({'rows':rows,'S':S,'g0':g0},open('meta.json','w'))
