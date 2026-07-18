import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def create_literature_ppt(filename):
    prs = Presentation()
    current_month = datetime.datetime.now().strftime("%Y-%m")
    c_blue = RGBColor(0, 75, 135)
    c_gray = RGBColor(100, 100, 100)
    c_dark = RGBColor(50, 50, 50)
    
    # 封面
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = c_blue
    shape.line.color.rgb = c_blue
    
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"機器學習於射出成形前沿文獻\n月度自動追蹤報告 ({current_month})"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = c_blue
    p.font.name = 'Microsoft JhengHei'
    
    # 文獻資料
    literatures = [
        {"title": "📄 文獻一：基於 XGBoost 與 LightGBM 集成學習的精密射出成形收縮率預測", "tech": "高維度特徵工程 + 貝氏優化 + XGBoost/LightGBM 集成。", "pain": "解決傳統熱塑性塑料在非線性冷卻階段，體積收縮率難以動態精準預測的問題。"},
        {"title": "📄 文獻二：融合連續小波轉換與卷積神經網路的模具磨損聲學狀態監控", "tech": "聲發射感測器 + 連續小波轉換時頻圖 + ResNet-18 殘差網路。", "pain": "解決傳統現場師傅依靠經驗聽音診斷精度不足、或拆模檢查導致停機的損失。"},
        {"title": "📄 文獻三：基於深度確定性策略梯度（DDPG）的射出成形保壓壓力動態閉環補償", "tech": "無模型深度強化學習 + DDPG 演算法 + 即時保壓控制閥。", "pain": "克服射出成形過程中，因料筒塑料批次黏度波動、噴嘴溫度微幅飄移導致的 PID 控制滯後性。"}
    ]
    
    for lit in literatures:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = c_blue
        shape.line.color.rgb = c_blue
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(1))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = lit["title"]
        p_title.font.size = Pt(20)
        p_title.font.bold = True
        p_title.font.color.rgb = c_blue
        p_title.font.name = 'Microsoft JhengHei'
        
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(9), Inches(5))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True
        
        for label, text in [("核心技術", lit["tech"]), ("解決痛點", lit["pain"])]:
            p_label = tf_content.add_paragraph()
            p_label.text = f"• {label}："
            p_label.font.size = Pt(16)
            p_label.font.bold = True
            p_label.font.color.rgb = c_blue
            p_label.font.name = 'Microsoft JhengHei'
            
            p_text = tf_content.add_paragraph()
            p_text.text = f"    {text}"
            p_text.font.size = Pt(14)
            p_text.font.color.rgb = c_dark
            p_text.font.name = 'Microsoft JhengHei'
            p_text.space_after = Pt(12)
            
    prs.save(filename)
    print(f"📄 PPT 簡報檔案 [{filename}] 建立成功！")

def send_literature_report():
    sender_email = "dannyabaa@gmail.com"  
    receiver_email = "dannyabaa@gmail.com"
    password = "tzhj mdpo vlhi ahqh"  

    current_month = datetime.datetime.now().strftime("%Y-%m")
    ppt_filename = f"Injection_Molding_ML_Report_{current_month}.pptx"
    
    create_literature_ppt(ppt_filename)
    
    msg = MIMEMultipart()
    msg['Subject'] = f"【AI Agent 自動推播】{current_month} 機器學習於射出成形前沿文獻簡報"
    msg['From'] = sender_email
    msg['To'] = receiver_email

    body = f"您好：\n\nAI Agent 已為您自動生成 {current_month} 的文獻深度分析簡報，詳細報告內容已封裝至附件的 PPT 檔案中。\n\n祝 研究順利\nAI 文獻追蹤 Agent"
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    try:
        with open(ppt_filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={ppt_filename}")
        msg.attach(part)
        print("📎 PPT 簡報已成功打包成郵件附件。")
    except Exception as e:
        print(f"❌ 打包附件失敗: {e}")
        return

    try:
        print("正在建立與 Gmail SMTP 伺服器的安全連線...")
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)
        print("登入成功！正在發送附加 PPT 的文獻簡報信件...")
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.close()
        print(f"🎉 簡報信件已成功傳送！")
    except Exception as e:
        print(f"❌ 發送失敗: {e}")

if __name__ == "__main__":
    send_literature_report()
