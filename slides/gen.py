import json, datetime
A=[0, -0.82, -1.92, -3.03, -4.13, -5.23, -6.33, -7.43, -9.74, -12.9, -15.67, -18.13, -20.31, -22.18, -23.86, -25.42, -26.78, -28.02, -29.25, -30.4, -31.48, -32.52, -33.56, -34.59]
Bv=[0, -0.39, -1.05, -1.72, -2.38, -3.05, -3.72, -4.38, -6.69, -9.84, -12.62, -15.08, -17.26, -19.13, -20.81, -22.37, -23.73, -24.97, -26.2, -27.35, -28.43, -29.47, -30.5, -31.54]
H="'Prompt', Tahoma, sans-serif"; Tx="'Noto Sans Thai', Tahoma, sans-serif"
DARK='#17332B'; LIGHT='#F6F4EE'; BODY='#44524D'; ACC='#A85A17'; GREEN='#245B4A'; CARD='#FDFCF9'; LINE='#E2DED3'; DBG='#1E3F35'; DBODY='#C9DCD5'; ORANGE='#E39A4B'
def sec(i,body,dark=False,foot=None,notes=''):
    bg=DBG if dark else LIGHT; col=LIGHT if dark else DARK
    pad='128px 128px 160px' if foot else '128px'
    f=f'<p style="position:absolute;left:128px;bottom:64px;width:1664px;font-size:24px;color:{DBODY if dark else "#6B7772"}">{foot}</p>' if foot else ''
    n=f'<aside>{notes}</aside>' if notes else ''
    return f'<section id="{i}" data-transition="fade" style="background:{bg};color:{col};font-family:{Tx};padding:{pad};display:flex;flex-direction:column;gap:40px">\n{body}\n{f}{n}</section>\n'
def head(eyebrow,title,dark=False):
    return f'<div style="display:flex;flex-direction:column;gap:12px"><p style="font-size:24px;font-weight:600;color:{ORANGE if dark else ACC};letter-spacing:1px">{eyebrow}</p><h2 style="font-family:{H};font-size:64px;font-weight:600;line-height:1.15;color:{LIGHT if dark else DARK}">{title}</h2></div>'
def card(title,text,flex=1,bg=CARD,tc=DARK,bc=BODY,border=LINE):
    return f'<div style="flex:{flex};display:flex;flex-direction:column;gap:12px;background:{bg};padding:36px;border:1px solid {border};border-radius:16px"><h3 style="font-family:{H};font-size:32px;font-weight:600;line-height:1.25;color:{tc}">{title}</h3><p style="font-size:24px;line-height:1.5;color:{bc}">{text}</p></div>'
def big(num,label):
    return f'<div style="flex:1;display:flex;flex-direction:column;gap:8px;background:{CARD};padding:36px;border:1px solid {LINE};border-radius:16px"><p style="font-family:{H};font-size:72px;font-weight:600;line-height:1.1;color:{GREEN}">{num}</p><p style="font-size:24px;line-height:1.45;color:{BODY}">{label}</p></div>'
def table(h,rows,fs=24,first_w=None):
    th=''.join(f'<th style="text-align:left;padding:14px 18px;background:#E6EEEA;color:{DARK};font-size:{fs}px;font-weight:600">{x}</th>' for x in h)
    tr=''.join('<tr>'+''.join(f'<td style="padding:14px 18px;border-bottom:1px solid {LINE};font-size:{fs}px;line-height:1.4;color:{DARK if j==0 else BODY}{";font-weight:600" if j==0 else ""}">{c}</td>' for j,c in enumerate(r))+'</tr>' for r in rows)
    return f'<table style="width:1664px;border-collapse:collapse;background:{CARD}"><tr>{th}</tr>{tr}</table>'
S={}
S['cover']=f'''<section id="cover" style="background:{DBG};color:{LIGHT};font-family:{Tx};padding:128px;display:flex;flex-direction:column;justify-content:space-between">
<p style="font-size:24px;font-weight:600;color:{ORANGE};letter-spacing:1px">บริษัท ซีเอ็มเอ็นเอช 2012 จำกัด · ร่างเพื่อตัดสินใจ</p>
<div style="display:flex;flex-direction:column;gap:24px"><h1 style="font-family:{H};font-size:110px;font-weight:600;line-height:1.1;color:{LIGHT}">แผนงานปี 2570</h1><p style="font-size:40px;line-height:1.4;color:{DBODY}">ครึ่งปีแรกจัดระบบ ครึ่งปีหลังเปิดโรงพยาบาลและ Villa</p></div>
<p style="font-size:24px;color:{DBODY}">ประชุม CEO และทีมผู้บริหาร · พฤศจิกายน 2569</p>
<aside>เปิดด้วยการบอกว่าวันนี้เป็นร่างเพื่อให้ตัดสินใจ ไม่ใช่แผนสุดท้าย ตัวเลขหลายตัวยังเป็นสมมติฐาน</aside>
</section>
'''
S['ask']=sec('ask',head('วันนี้','ขอมติจากที่ประชุม 3 เรื่อง')+'<div style="display:flex;gap:32px">'+card('1 · เลือกทางเลือก A หรือ B','ทำทุกอย่างในปี 2570 หรือเลื่อนตลาดต่างประเทศและเปิด OPD เป็นขั้น')+card('2 · อนุมัติโครงสร้างใหม่','3 สาย 8 แผนก ควบ 4 ตำแหน่ง (JD 21) และมอบอำนาจอนุมัติให้หัวหน้าแผนกและ Vice CEO')+card('3 · มอบหมายเก็บตัวเลขจริง','ดอกเบี้ยเงินกู้จริง ประมาณการ OPD และแพ็กเกจ ภายใน 23 พ.ย. 69 และแก้คำขอใบอนุญาต รพ. ก่อนยื่น')+'</div>',notes='บอกตั้งแต่ต้นว่าต้องการอะไรจากที่ประชุม แล้วค่อยลงรายละเอียด')
S['year']=sec('year',head('ภาพรวม','ปี 2570 มี 4 เรื่องหลัก')+'<div style="display:grid;grid-template-columns:1fr 1fr;gap:28px">'+card('คลินิกเวชกรรม','เปิดที่โรงพยาบาล CMNH ใน Q1')+card('โครงสร้างใหม่ + KPI','ปรับใน Q1 ประเมินรอบแรกใน Q2')+card('โรงพยาบาลและ Villa','วางโครงสร้าง รพ. Q3 · เปิด Villa Q4')+card('ตลาดต่างประเทศ','วางรากฐาน Q1–Q2 กับ Swiss Helping Point · เริ่มขายในสวิส Q3 (B: ทดลองขาย Q4)')+'</div>',foot='งานต่อเนื่องทั้งปี: ระบบแฟรนไชส์กับ Franbiz · ติดตามก่อสร้าง รพ. กับ DG และธนาคาร')
def qcol(t,m,items,hl=None):
    it=''.join(f'<p style="font-size:24px;line-height:1.4;color:{DARK if x==hl else BODY};{"font-weight:600;" if x==hl else ""}padding:12px 16px;background:{"#E6EEEA" if x==hl else "#FFFFFF"};border:1px solid {GREEN if x==hl else LINE};border-radius:10px">{x}</p>' for x in items)
    return f'<div style="flex:1;display:flex;flex-direction:column;gap:12px"><h3 style="font-family:{H};font-size:32px;font-weight:600;color:{DARK}">{t}</h3><p style="font-size:24px;color:#6B7772">{m}</p>{it}</div>'
S['quarters']=sec('quarters',head('แผนรายไตรมาส','ครึ่งปีแรกจัดระบบ ครึ่งปีหลังขยาย')+'<div style="display:flex;gap:24px">'+qcol('Q1 · ตั้งระบบ','ม.ค.–มี.ค.',['แก้คำขอ ส.พ. · ยื่นคลินิกแยก','เปิดคลินิกเวชกรรม','รูปแบบการจัดการใหม่','ยุทธศาสตร์ + ผังใหม่','กำหนด KPI','วางรากฐาน Tranquila Intl.'])+qcol('Q2 · วัดผล','เม.ย.–มิ.ย.',['ประเมิน KPI','ประเมิน Performance','ทบทวนแผนกลางปี','เริ่มสรรหาคน รพ.'],'เริ่มสรรหาคน รพ.')+qcol('Q3 · วางแผนขยาย','ก.ค.–ก.ย.',['วางโครงสร้าง รพ.','เปิดตัว Tranquila Intl. ในสวิส','ยื่นขออนุญาต Villa'],'วางโครงสร้าง รพ.')+qcol('Q4 · เริ่มดำเนินการ','ต.ค.–ธ.ค.',['เปิด IPD (สมมติ ต.ค.)','เปิด Villa ใหม่','ผู้พักต่างชาติทดลองชุดแรก','สรุปผลปี + แผน 2571'],'เปิด Villa ใหม่')+'</div>',notes='ทุกงานมีขั้นตอนย่อยและเอกสารประกอบใน CEO Dashboard')
def orgcol(t,deps):
    d=''.join(f'<p style="font-size:24px;line-height:1.4;color:{BODY}"><b style="color:{DARK}">{a}</b><br>{b}</p>' for a,b in deps)
    return f'<div style="flex:1;display:flex;flex-direction:column;gap:16px;background:{CARD};padding:32px;border:1px solid {LINE};border-radius:16px"><h3 style="font-family:{H};font-size:32px;font-weight:600;color:{GREEN}">{t}</h3>{d}</div>'
S['org']=sec('org',head('โครงสร้างองค์กร','3 สาย 8 แผนก JD 21 ตำแหน่ง (เสนอควบ 4 คู่)')+'<div style="display:flex;gap:24px">'+orgcol('COO',[('พยาบาล','Head of Staff, BM, RN, PN, NA, คนขับรถ'),('กายภาพ · กิจกรรมบำบัด','PT, PTA, OT'),('บริการ · บุคคลและธุรการ','แม่ครัว, แม่บ้าน, HR & Admin')])+orgcol('CFO',[('บัญชีและการเงิน','Finance & Accounting'),('จัดซื้อและคลังสินค้า','Purchasing & Inventory')])+orgcol('CMO · การตลาดและกลยุทธ์',[('การตลาด','Marketing & Sales'),('รวมงาน CSO เดิม','พันธมิตร · ตลาดต่างประเทศ')])+'</div>',foot='ควบ 4 คู่: CMO+CSO · จัดซื้อ+คลัง · HR+ธุรการ · คนขับรถ+เวรเปล (ไม่ลดคน)',notes='JD ฉบับเต็มทั้ง 21 ตำแหน่งอยู่ใน Dashboard แท็บโครงสร้าง')
S['ceo']=f'''<section id="ceo" data-transition="fade" style="background:{DBG};color:{LIGHT};font-family:{Tx};padding:128px;display:flex;flex-direction:column;justify-content:center;gap:40px">
<p style="font-size:24px;font-weight:600;color:{ORANGE};letter-spacing:1px">ข้อสังเกตสำคัญ</p>
<h2 style="font-family:{H};font-size:80px;font-weight:600;line-height:1.2;color:{LIGHT}">ปี 2570 CEO ดำรงตำแหน่ง CEO, COO, CFO และ CMO พร้อมกัน</h2>
<p style="font-size:40px;line-height:1.45;color:{DBODY}">เสนอให้หัวหน้าแผนกรับ KPI และอำนาจอนุมัติตามวงเงิน และให้ Vice CEO รับโครงการโรงพยาบาล Villa และแฟรนไชส์ เพื่อไม่ให้การตัดสินใจติดคอขวด</p>
<aside>นี่คือเหตุผลหลักที่เสนอทางเลือก B: ลดจำนวนโครงการที่ต้องดูพร้อมกัน</aside>
</section>
'''
S['hosp']=sec('hosp',head('รพ. เชียงใหม่ มาสเตอร์พีซ (CMNH) + Villa','ตัวเลขเมื่อเปิดเต็ม ตามบุคลากรในคำขอใบอนุญาต')+'<div style="display:flex;gap:28px">'+big('70','เตียง: IPD 50 · Villa 20')+big('118','พนักงาน: OPD 31 · IPD 68 · Villa 19')+big('4.53 ล.','ค่าใช้จ่าย/เดือน รวมค่าผ่อน 0.90 ล.')+big('1.04 ล.','รายได้ OPD + แพ็กเกจ/เดือนที่ต้องมีเพิ่ม')+'</div><p style="font-size:30px;line-height:1.5;color:'+BODY+'">ต้นทุนเฉลี่ย 51,843 บาท/คน/เดือน สูงกว่าราคาโซนรวม 45,000 บาท ค่าเตียงอย่างเดียวไม่คุ้มทุนแม้เข้าพักเต็ม ต้องมีรายได้ OPD และแพ็กเกจเสริม</p>',foot='ที่มา: CMNH_แผนการเงิน_2570-71.xlsx · ส.พ.4/ส.พ.6 · ลงทุน 161.8 ล้าน กู้ 66% 6% 15 ปี (ดอกเบี้ยสมมติ)',notes='ปรับบุคลากรตามคำขอใบอนุญาต และใส่ค่าผ่อนธนาคาร')
# chart svg
W,Hh=1664,440; pl,pr,pt,pb=90,30,20,50
mn,mx=-42,0
X=lambda i:pl+i*(W-pl-pr)/23; Y=lambda v:pt+(mx-v)*(Hh-pt-pb)/(mx-mn)
svg=f'<svg aria-label="เงินสดสะสม 24 เดือน ทางเลือก A ลบ 34.59 ล้านบาท ทางเลือก B ลบ 31.54 ล้านบาท ที่เดือนที่ 24" width="{W}" height="{Hh}" viewBox="0 0 {W} {Hh}">'
for t in range(0,-43,-6):
    svg+=f'<line x1="{pl}" x2="{W-pr}" y1="{Y(t):.0f}" y2="{Y(t):.0f}" stroke="{"#9AA5A1" if t==0 else "#E2DED3"}" stroke-width="1.5"/><text x="{pl-14}" y="{Y(t)+8:.0f}" font-size="24" text-anchor="end" fill="#6B7772">{t}</text>'
MON=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
for i in range(0,24,3):
    svg+=f'<text x="{X(i):.0f}" y="{Hh-12}" font-size="24" text-anchor="middle" fill="#6B7772">{MON[i%12]} {70+i//12}</text>'
for vals,col,lab,dy in [(A,'#A85A17','ทางเลือก A −34.6 ล้าน',38),(Bv,GREEN,'ทางเลือก B −31.5 ล้าน',-42)]:
    svg+=f'<polyline fill="none" stroke="{col}" stroke-width="5" points="{" ".join(f"{X(i):.0f},{Y(v):.0f}" for i,v in enumerate(vals))}"/>'
    svg+=f'<circle cx="{X(23):.0f}" cy="{Y(vals[23]):.0f}" r="9" fill="{col}"/><text x="{X(23)-18:.0f}" y="{Y(vals[23])+dy:.0f}" font-size="26" font-weight="600" text-anchor="end" fill="{col}">{lab}</text>'
svg+='</svg>'
S['cash']=sec('cash',head('แผนการเงินรายเดือน','ไม่มีรายได้ OPD และแพ็กเกจ เงินสดติดลบ 31.5–34.6 ล้าน')+svg,foot='เงินสดสะสม (ล้านบาท) · รวมบุคลากรตามใบอนุญาตและค่าผ่อนธนาคาร (เริ่ม ต.ค. 70) · ยังไม่รวมรายได้ OPD และแพ็กเกจเสริม',notes='สมมติฐาน: เปิด OPD มี.ค. 70 · IPD ต.ค. 70 · Villa พ.ย. 70 · เข้าพักเริ่ม 15% ไปถึง 85% ใน 12 เดือน ทดลองปรับได้ในแท็บการเงินของ Dashboard')
S['options']=sec('options',head('ทางเลือกให้ตัดสินใจ','เสนอทางเลือก B: ประหยัดเงินราว 3 ล้านบาท และลดงานผู้บริหาร')+table(['เรื่อง','A: ทำทั้งหมดในปี 2570','B: เลื่อนตลาดต่างประเทศ + เปิด OPD เป็นขั้น'],[['OPD ก่อนเปิด IPD','เปิดเต็ม 2 ห้องตรวจ · 822,000 บาท/เดือน','1 ห้องตรวจ 08.00–17.00 · 386,000 บาท/เดือน'],['ตลาดต่างประเทศ','วางรากฐาน Q1–Q2 เริ่มขายในสวิส ก.ค. 70','วางรากฐาน Q1–Q2 ทดลองขาย Q4 ขายเต็มปี 2571'],['เงินสดสะสม 24 เดือน','−34.59 ล้านบาท','<b style="color:'+GREEN+'">−31.54 ล้านบาท</b>'],['ผลปี 2570','−18.13 ล้านบาท','−15.08 ล้านบาท'],['เดือนแรกที่มีกำไร','เกิน 24 เดือน*','เกิน 24 เดือน*'],['โครงการที่ดูพร้อมกัน','7','6']],fs=26),notes='* ถ้าไม่มีรายได้ OPD และแพ็กเกจ ทั้งสองทางยังไม่มีกำไรใน 24 เดือน B ช่วยประหยัดเงินช่วงก่อนเปิดราว 3.05 ล้านบาท')
S['sales']=sec('sales',head('แผนการขาย','ต้องหาผู้สนใจ 96 รายในปี 2570 และ 232 รายในปี 2571')+table(['แหล่งผู้พัก (ข้อเสนอ)','สัดส่วน','ปี 2570','ปี 2571','สายที่ดูแล'],[['ส่งต่อจากโรงพยาบาลพันธมิตร','35%','34','81','CMO'],['คลินิกหมอต้น / OPD','20%','19','46','COO'],['ออนไลน์และการตลาด','20%','19','46','CMO'],['ย้ายจากสาขา 1–4 และ TQL','15%','14','35','COO'],['ตลาดต่างประเทศ','10%','10','23','CMO + CSO']]),foot='สมมติ: ปิดการขาย 25% · ผู้พักออก 4%/เดือน · เริ่มหาผู้สนใจก่อนเปิดอย่างน้อย 2 เดือน')
S['risks']=sec('risks',head('ความเสี่ยงและเกณฑ์ก่อนเปิด','แต่ละหน่วยเปิดเมื่อผ่านทุกข้อ ไม่ผ่านให้เลื่อน')+'<div style="display:flex;gap:28px">'+card('คลินิกเวชกรรม (มี.ค. 70)','ใบอนุญาต สสจ. · แพทย์ผู้ดำเนินการ · ทีม OPD ครบ · ประมาณการรายได้ที่ CFO รับรอง')+card('IPD 50 เตียง (ต.ค. 70)','ใบอนุญาตรับผู้ป่วยค้างคืน · อาคารผ่านตรวจรับ · บุคลากรครบตาม ส.พ.6 (แพทย์ RN เภสัชกร MT RT OT) · แพทย์เวรกลางคืน · จองล่วงหน้า 8 คน · เงินสำรองครบ')+card('Villa 20 ห้อง (พ.ย. 70)','ก่อสร้างเสร็จ · ใบอนุญาตถูกประเภท · care manager + 3 คน · จอง 3 ห้อง · ราคายืนยันแล้ว')+'</div><p style="font-size:30px;line-height:1.5;color:'+BODY+'">ความเสี่ยงสูงสุด: <b style="color:'+DARK+'">หา RN ไม่ทัน</b> จึงเริ่มสรรหาตั้งแต่ Q2 และ <b style="color:'+DARK+'">ใบอนุญาตผิดประเภท</b> จึงยืนยันกับหน่วยงานใน Q1</p>')
S['next']=sec('next',head('ก้าวต่อไป','งานเร่งด่วนก่อนเริ่มปี')+table(['เมื่อไร','งาน','ใคร'],[['ภายใน 23 พ.ย. 69','เก็บตัวเลขจริง: ธุรกิจเดิมปี 2569 · เงื่อนไขเงินกู้ · ประมาณการ OPD และแพ็กเกจ · วันส่งมอบอาคาร · ราคา Villa · แก้คำขอใบอนุญาต รพ. (ผู้ดำเนินการ แพทย์เวร เอกสารบุคลากร)','CFO + COO + CMO + Vice CEO'],['ธ.ค. 69','ยื่นขอใบอนุญาตคลินิกเวชกรรม · เซ็นสัญญาแพทย์ผู้ดำเนินการ','COO'],['ม.ค. 70','กรรมการอนุมัติยุทธศาสตร์ · Town hall · ยืนยันประเภทใบอนุญาต Villa','CEO + Vice CEO'],['ม.ค. 70','ตั้งค่าเป้า KPI ทุกสาย จากตัวเลขฐานปี 2569','ทุกสาย (HR ประสาน)']],fs=26))
S['close']=f'''<section id="close" data-transition="fade" style="background:{DBG};color:{LIGHT};font-family:{Tx};padding:128px;display:flex;flex-direction:column;justify-content:center;gap:48px">
<h2 style="font-family:{H};font-size:80px;font-weight:600;line-height:1.2;color:{LIGHT}">ขอมติวันนี้</h2>
<div style="display:flex;flex-direction:column;gap:24px">
<p style="font-size:40px;line-height:1.4;color:{DBODY}"><span style="color:{ORANGE};font-weight:600">1</span>  เลือกทางเลือก A หรือ B</p>
<p style="font-size:40px;line-height:1.4;color:{DBODY}"><span style="color:{ORANGE};font-weight:600">2</span>  อนุมัติโครงสร้าง 8 แผนก การควบตำแหน่ง และการมอบอำนาจ</p>
<p style="font-size:40px;line-height:1.4;color:{DBODY}"><span style="color:{ORANGE};font-weight:600">3</span>  มอบหมายเก็บตัวเลขจริง 6 รายการ ภายใน 23 พ.ย. 69</p>
</div>
<p style="font-size:24px;color:{DBODY}">รายละเอียดทั้งหมด: เอกสารแผนงานปี 2570 · CEO Dashboard · CMNH_แผนการเงิน_2570-71.xlsx</p>
</section>
'''
order=['cover','ask','year','quarters','org','ceo','hosp','cash','options','sales','risks','next','close']
for k in order: open(f'project/slides/{k}.html','w').write(S[k].strip()+'\n')
deck={"v":4,"createdOnFiles":{"v":1,"at":datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')},"lists":"css","title":"CMNH แผนงานปี 2570 — คุยกับ CEO และทีม","order":order,
"sections":{"s1":{"description":"เปิดด้วยมติที่ต้องการ","start":"cover"},"s2":{"description":"แผนงานและโครงสร้าง","start":"year"},"s3":{"description":"ตัวเลขโรงพยาบาลและทางเลือก","start":"hosp"},"s4":{"description":"การขาย ความเสี่ยง และก้าวต่อไป","start":"sales"}},
"cover":"cover","faces":{"prompt":{"family":"Prompt","href":"https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap"},"noto-sans-thai":{"family":"Noto Sans Thai","href":"https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400..700&display=swap"}},"designSystems":[]}
json.dump(deck,open('project/deck.json','w'),ensure_ascii=False,indent=1)
print('ok',len(order))
