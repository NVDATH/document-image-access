# ชุดไฟล์ทดสอบ docImages 0.6.1

ภาพทุกภาพมีรูปทรงและข้อความกำกับ เช่น "RED SQUARE" เพื่อให้ Be My Eyes อธิบายแล้วตรวจได้ว่าส่งถูกใบ ไฟล์ 1, 2 และ 5 ปรับใหม่ให้ชาร์ตมีชื่อ (title) และ alt text

## ผลที่ควรได้ยินเมื่อกดครั้งเดียว (NVDA+Shift+I)

| ไฟล์ | ประกาศ |
|---|---|
| 1a_excel_floating_pictures_and_chart.xlsx | 2 images, all with alt text. 1 chart, which is not an image. |
| 1b_excel_header_picture_only.xlsx | 1 image, 1 without alt text |
| 1c_excel_background_only.xlsx | 1 image, 1 without alt text |
| 1d_excel_in_cell_pictures_only.xlsx | 2 images, 1 without alt text |
| 2_excel_chart_only_no_pictures.xlsx | No images found in 2_excel_chart_only_no_pictures.xlsx. 1 chart, which is not an image. |
| 3_excel_picture_in_shape_fill.xlsx | 1 image, all with alt text. 1 picture file in the document could not be located. |
| 4_word_body_table_header_footer.docx | 5 images, 4 without alt text |
| 4b_word_with_chart.docx | 1 image, 1 without alt text. 1 chart, which is not an image. |
| 5_powerpoint_pictures_and_chart.pptx | 3 images, 2 without alt text. 1 chart, which is not an image. |

## รายการเต็มที่ควรเห็น (กดซ้ำสองครั้งไวๆ)

ใหม่ใน 0.5.0: ชาร์ตของ Excel และ PowerPoint อยู่ในรายการด้วย ต่อท้ายรูป ไฟล์ที่มีแต่ชาร์ต (ไฟล์ 2) ก็เปิดรายการได้

ไฟล์ 1a มี 3 แถว:
- Image 1; Sheet Pictures, cell B3; Alt text: Red square with a black outline
- Image 2; Sheet Pictures, cell F3; Alt text: Blue circle
- Chart 1; Sheet Pictures, cell H12; Alt text: Column chart of quarterly sales; Title: Quarterly sales

ไฟล์ 1b: Image 1; Sheet Header, left header; No alt text
ไฟล์ 1c: Image 1; Sheet Background, background; No alt text
ไฟล์ 1d: Image 1; Sheet InCell, cell A1, in cell; Alt text: Green triangle in a cell และ Image 2; Sheet InCell, cell A3, in cell; No alt text

ไฟล์ 2 มี 1 แถว: Chart 1; Chart sheet Chart1; No alt text; Title: Growth over time

ไฟล์ 5 มี 4 แถว: ภาพ 3 รูป (Slide 1, Slide 2, Slide 2) แล้วต่อด้วย Chart 1; Slide 2; Alt text: Column chart of results by group; Title: Results by group

ไฟล์ 4b (Word มีชาร์ต): 2 แถว คือ Image 1 (Near: Picture after the chart:) และ Chart 1 (Alt text: Column chart of sales by region; Title: Sales by region) ถ้ามี Word ตำแหน่งต้องเป็นเลขหน้าและบรรทัดจริง

ไฟล์ 4 (Word): 5 แถว ถ้ามี Word ตำแหน่งต้องเป็นเลขหน้าและบรรทัดจริง ภาพ header และ footer ต้องบอก "In header" และ "In footer"

## ไฟล์ 1d เปิดใน Excel ไม่ได้
ผลทดสอบของคุณ: 1a, 1b, 1c เปิดได้ ส่วน 1d (รูปในเซลล์ Place in Cell) เปิดไม่ได้ใน Office ของคุณ ผมสงสัยว่า Office รุ่นนั้นไม่รองรับรูปในเซลล์ (จากความจำของผม ฟีเจอร์นี้เป็นของ Microsoft 365) แต่ยังไม่ได้ยืนยัน ไฟล์ 1d ใช้ทดสอบจาก File Explorer ได้ตามปกติ ไม่ต้องเปิดใน Excel

## ทดสอบฟีเจอร์ที่ต้องเปิดไฟล์ในโปรแกรมจริง (กด Enable Editing ก่อน)

เปิด 1a ใน Excel แล้วกดคีย์ลัดสองครั้งไวๆ:
1. Image 1 กด Go to in document: เคอร์เซอร์ควรไปที่เซลล์ B3
2. Image 2 กด Edit alt text: ช่องควรแสดงข้อความ alt ปัจจุบัน เปลี่ยนแล้วตรวจใน Excel
3. Chart 1 กด Edit alt text: ควรเปลี่ยนได้เหมือนรูป ถ้าไม่ได้ ให้จดข้อความที่ได้ยิน
4. Chart 1 กด Open หรือ Send to Be My Eyes: ควรส่งออกชาร์ตแล้วเปิดหรือส่งได้
5. ลองก่อนกด Enable Editing: Chart 1 กด Open ควรได้ยิน "The document is open in Protected View. Choose Enable Editing..."
6. หลังกดปุ่มใดๆ (Open, Save as, Copy image, Send to Be My Eyes, Edit alt text, Export all) ข้อความผลลัพธ์ (เช่น "Alt text updated.") ต้องพูดก่อน แล้วโฟกัสจึงกลับมาที่รายการ ถ้าข้อความถูกประกาศของ NVDA ตัดจนฟังไม่ทัน ให้แจ้งผม
7. จุดโฟกัสแถว: ย้ายเซลล์ไปใกล้ F3 แล้วเปิดรายการ ควรเริ่มที่ Image 2 ย้ายไปใกล้ H12 ควรเริ่มที่ Chart 1 กด Go to แล้วเปิดรายการใหม่ ควรอยู่แถวเดิม

เปิดไฟล์ 2 ใน Excel: Chart 1 กด Go to จะเปิดชีตชาร์ต Chart1

เปิดไฟล์ 5 ใน PowerPoint: เปิดบนสไลด์ 2 แล้วเปิดรายการ ควรเริ่มที่รูปของสไลด์ 2 (ถ้าเลือกรูปไว้ ควรเริ่มที่รูปนั้น) Go to, Edit alt text และส่งออกชาร์ตควรทำงาน

เปิดไฟล์ 4 ใน Word: เลื่อนเคอร์เซอร์ไปใกล้ภาพที่สอง แล้วเปิดรายการ ควรเริ่มที่ภาพนั้น

เปิดไฟล์จาก Explorer แล้วเปิดรายการ: Go to และ Edit alt text ใช้ไม่ได้ ปิดรายการแล้วเปิดใหม่ ควรอยู่แถวเดิม

## ทดสอบใหม่ใน 0.6.0

**ชาร์ตของ Word** (ไฟล์ 4b เปิดใน Word กด Enable Editing ก่อน): เปิดรายการ ลองกับแถว Chart 1
1. Go to in document: เคอร์เซอร์ควรไปที่ชาร์ต
2. Edit alt text: ช่องควรแสดง "Column chart of sales by region" เปลี่ยนแล้วตรวจใน Word
3. Open หรือ Send to Be My Eyes: ควรส่งออกชาร์ตเป็นรูปแล้วเปิดหรือส่งได้ (Word ส่งออกด้วย Chart.Export ถ้าไม่ทำงาน log จะบอกสาเหตุ)

**ปุ่ม Whole page or slide** (ส่งออกทั้งหน้าหรือทั้งสไลด์เป็นรูป เพื่อให้ Be My Eyes อธิบายภาพรวม) กดแล้วจะมีเมนู: Open, Save as, Copy to the clipboard, Send to Be My Eyes
- PowerPoint (ไฟล์ 5): เลือก Image 2 แล้วกดปุ่มนี้ ควรได้รูปทั้งสไลด์ 2
- Word (ไฟล์ 4 หรือ 4b): เลือกรูปแล้วกดปุ่มนี้ ควรได้รูปเนื้อหาทั้งหน้าที่รูปนั้นอยู่ วิธีของ Word คือสั่งให้ Word วาดเนื้อหาของหน้าเป็นภาพ (metafile) ซึ่งผมไม่แน่ใจว่าหน้าตาจะเหมือนหน้าจริงแค่ไหน ถ้าผลเพี้ยน บอกผมได้
- Excel: ควรได้ยินว่า "Exporting a whole sheet is not available. Only Word pages and PowerPoint slides can be exported."
- ใน Word ปุ่มนี้ใช้ได้เมื่อรู้เลขหน้าจริง ถ้าอ่านเลขหน้าจาก Word ไม่สำเร็จ จะได้ยินว่าไม่รู้หน้า

## ทดสอบการตรวจ Be My Eyes ตอนติดตั้ง (0.6.1)

เครื่องที่มี Be My Eyes อยู่แล้ว: ติดตั้งส่วนเสริมแล้วไม่ควรมีหน้าต่างเด้งขึ้นมา ตรวจว่ามีไฟล์ docImages.json ในโฟลเดอร์ตั้งค่าของ NVDA และมีบรรทัด "aumid" อยู่ (ทำให้การส่งครั้งแรกไม่ต้องรอค้นหา)

เครื่องที่ไม่มี (ถอนการติดตั้ง Be My Eyes ชั่วคราวเพื่อทดสอบ) ติดตั้งส่วนเสริมแล้วควรมีหน้าต่าง "Be My Eyes not found" พร้อมปุ่ม:
1. Open the Microsoft Store to install Be My Eyes: เปิดหน้า Be My Eyes ในร้านค้า (ถ้าเปิดไม่ได้ จะบอกให้ใช้ปุ่ม Copy)
2. Copy the Microsoft Store link: คัดลอกลิงก์ไว้วางในเบราว์เซอร์
3. Check again: ตรวจใหม่หลังติดตั้งเสร็จ ถ้าเจอ หน้าต่างจะปิดและบอกว่าพบแล้ว
4. ช่องติ๊ก Do not remind me again แล้วกด Continue without Be My Eyes: ครั้งหน้าที่ติดตั้งหรืออัปเกรด จะไม่ถามอีก

ตอนใช้งาน: บนเครื่องที่ไม่มี Be My Eyes กด Send to Be My Eyes แล้วควรมีหน้าต่างถามว่าจะเปิดหน้าร้านค้าไหม ตอบ Yes ควรเปิดร้านค้า

## ข้อควรรู้เพิ่ม
- ไฟล์ 4b ผมประกอบเอง โดยนำชาร์ตที่ไลบรารี python-pptx สร้างมาใส่ใน docx ไม่ได้ลองเปิดใน Word ถ้า Word ขอซ่อมไฟล์หรือเปิดไม่ได้ ให้ทดสอบจาก File Explorer แทน ส่วนเสริมอ่านไฟล์ได้
## ข้อควรรู้
- ไฟล์ทั้งหมดสร้างด้วยไลบรารี (XlsxWriter, python-docx, python-pptx) ไม่ได้สร้างจากโปรแกรม Office ตัวจริง
- ไฟล์ 3 ผมแก้โครงสร้างภายในเอง ไม่ได้ลองเปิดใน Excel ถ้า Excel ขอซ่อมไฟล์ ให้ข้ามการเปิดใน Excel การทดสอบใช้แค่เลือกไฟล์ใน Explorer
- ส่วนที่เรียกโปรแกรม Office จริง ผมทดสอบกับวัตถุจำลองเท่านั้น คุณยืนยันแล้วว่าส่งชาร์ตไป Be My Eyes ใช้ได้ ส่วนที่เหลือรอผลทดสอบ ถ้าไม่ทำงาน ส่ง log ที่ค้นคำว่า docImages มา
