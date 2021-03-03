import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

mail = outlook.CreateItem(0)

mail.to = 'adorder@clea-japan.com'
mail.cc = 'nishimoto.asako.0v@kyoto-u.ac.jp;okamurah@pharm.kyoto-u.ac.jp;doimasao@pharm.kyoto-u.ac.jp'
mail.bcc = 'doi.shoutarou.44a@st.kyoto-u.ac.jp'
mail.subject = 'マウス餌CE-2（京大薬システムバイオ）'
mail.bodyFormat = 1
mail.body = '''日本クレア株式会社担当者さま 
 
いつもお世話になっております。
 
2月4日（木）の便で、下記の商品を注文致します。
                
品名：CREA Rodent Diet CE-2 数量:4袋(80kg）
 
すべて地下にお願いいたします。
 
土井 星太朗
<<<<<<<<<<<<<<<<<<<<<<<<<<<< 
土井 星太朗
京都大学 薬学部薬科学科 4回生 
医薬創成情報科学講座 
システムバイオロジー分野 
〒606-8501 京都市左京区吉田下阿達町46-29 別館4F 
Tel: 075-753-9554　Fax: 075-753-9553
Email: doi.shoutarou.44a@st.kyoto-u.ac.jp

'''

mail.display(True)