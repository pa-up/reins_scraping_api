from fastapi import FastAPI
from pydantic import BaseModel
import time
import os

import excel_or_csv as ec
from scraping import Reins_Scraper
import py_mail
from aws import ManipulateS3



static_path = "static"
# 入力パス
mail_excel_path = static_path + "/input_excel/email_pw.xlsx"
search_method_excel_path = static_path + "/input_excel/search_method.xlsx"
input_reins_excel_path = static_path + "/input_excel/input_reins.xlsx"

# 出力パス
output_reins_excel_path = static_path + "/output_excel/output_reins.xlsx"
log_txt_path = static_path + "/log/log.txt"

# 環境変数の取得
user_id , password = os.environ.get('SECRET_USER_ID') , os.environ.get('SECRET_PASSWORD')
s3_accesskey , s3_secretkey = os.environ.get('S3_ACCESSKEY') , os.environ.get('S3_SECRETKEY')
s3_bucket_name = os.environ.get('S3_BUCKET_NAME')


# s3の定義
s3_region = "ap-northeast-1"   # 東京(アジアパシフィック)：ap-northeast-1


class logText:
    def __init__(self , log_txt_path) -> None:
        self.log_txt_path = log_txt_path
        # logの保存ファイルを空にする
        with open(self.log_txt_path, 'w') as file:
            file.write('')

    def add_log_txt(self , add_log_text):
        """ logを付け加える関数 """
        with open(self.log_txt_path, 'a') as file:
            file.write("\n" + add_log_text)
log_txt = logText(log_txt_path)


class RequestDataScraping(BaseModel):
    search_method_value: list
    index_of_search_requirement: int
    mail_list: list
    cc_mail_list: list
    from_email: str
    from_email_smtp_password: str

class RequestDataExcel(BaseModel):
    many_excel_list: list
    search_method_list: str
    search_requirement_list: str
    mail_list: list
    cc_mail_list: list
    from_email: str
    from_email_smtp_password: str

# app = FastAPI()
app = FastAPI(default_response_limit=1024 * 1024 * 10)  # 10MBに増量


@app.post("/")
def fast_api_scraping():
    # S3からメール情報や検索条件を取得し、静的フォルダに格納
    manipulate_s3 = ManipulateS3(
        region = "ap-northeast-1" ,
        accesskey = s3_accesskey ,
        secretkey = s3_secretkey ,
        bucket_name = s3_bucket_name ,
    )
    manipulate_s3.s3_file_download(local_upload_path = mail_excel_path)
    manipulate_s3.s3_file_download(local_upload_path = search_method_excel_path)
    time.sleep(2)

    # Excelファイルから検索方法と検索条件を選択（別のWEBアプリでも編集可能）
    index_of_solding_requirement_list, index_of_rental_requirement_list = ec.get_search_option_from_excel(search_method_excel_path)
    index_of_solding_requirement_list = [x for x in index_of_solding_requirement_list if x not in (0, None)]
    index_of_rental_requirement_list = [x for x in index_of_rental_requirement_list if x not in (0, None)]


    # メールアドレスのリストをExcelから取得
    mail_list , cc_mail_list , from_email , from_email_smtp_password = ec.mail_list_from_excel(mail_excel_path)

    try:
        log_txt.add_log_txt("S3 取得完了")
        # ページにアクセス
        login_url = "https://system.reins.jp/"
        searched_url = "https://system.reins.jp/main/KG/GKG003100"
        reins_sraper = Reins_Scraper()
        
        log_txt.add_log_txt("reinsサイトにアクセス完了")

        # ログイン突破
        reins_sraper.login_reins(login_url , user_id , password)
        log_txt.add_log_txt("ログイン成功")
        # REINS上で存在する検索方法と検索条件を全て取得（01〜50番号まであることを前提）
        solding_search_method_list , rental_search_method_list = reins_sraper.get_solding_or_rental_option()
        log_txt.add_log_txt(f"solding_search_method_list , rental_search_method_list : \n{solding_search_method_list} \n{rental_search_method_list}")
        
        # スクレイピング結果のリストを取得
        many_excel_list = []
        search_method_list = []
        search_requirement_list = []
        if index_of_solding_requirement_list != []:
            log_txt.add_log_txt("条件 : index_of_solding_requirement_list != []:")
            for index_of_solding_requirement in index_of_solding_requirement_list:
                search_method_value = "search_solding"
                index_of_search_requirement = index_of_solding_requirement
                
                # 検索条件の名前に"土地"があれば、スクレイピングを実行
                search_requirement = solding_search_method_list[index_of_solding_requirement]
                if "土地" in search_requirement:
                    to_excel_list = reins_sraper.scraping_solding_list(searched_url , search_method_value , index_of_search_requirement)
                    print(f"スクレイピング後 : search_method_value : {search_method_value}")
                    if len(to_excel_list) == 1:
                        to_excel_list[0][0] = to_excel_list[0][0] + f"「{solding_search_method_list[index_of_solding_requirement]}」"
                else:
                    to_excel_list = [[ "その検索条件の名前に「土地」が入っていません : " + f"「{solding_search_method_list[index_of_solding_requirement]}」" ]]

                many_excel_list.append(to_excel_list)
                search_method_list.append(search_method_value)
                search_requirement_list.append(solding_search_method_list[index_of_solding_requirement])
        

        if index_of_rental_requirement_list != []:
            log_txt.add_log_txt("条件 : index_of_rental_requirement_list != []:")
            for index_of_rental_requirement in index_of_rental_requirement_list:
                search_method_value = "search_rental"
                index_of_search_requirement = index_of_rental_requirement

                # 検索条件の名前に"土地"があれば、スクレイピングを実行
                search_requirement = rental_search_method_list[index_of_rental_requirement]
                if "土地" in search_requirement:
                    to_excel_list = reins_sraper.scraping_solding_list(searched_url , search_method_value , index_of_search_requirement)
                    print(f"スクレイピング後 : search_method_value : {search_method_value}")
                    if len(to_excel_list) == 1:
                        to_excel_list[0][0] = to_excel_list[0][0] + f"「{rental_search_method_list[index_of_rental_requirement]}」"
                else:
                    to_excel_list = [[ "その検索条件の名前に「土地」が入っていません : " + f"「{rental_search_method_list[index_of_rental_requirement]}」" ]]

                many_excel_list.append(to_excel_list)
                search_method_list.append(search_method_value)
                search_requirement_list.append(rental_search_method_list[index_of_rental_requirement])
        
        reins_sraper.driver.quit()
        log_txt.add_log_txt("スクレイピング結果のリストを取得 : 完了")

        return {
            "many_excel_list": many_excel_list ,
            "search_method_list" : search_method_list ,
            "search_requirement_list" : search_requirement_list ,
            "mail_list" : mail_list ,
            "cc_mail_list" : cc_mail_list ,
            "from_email" : from_email ,
            "from_email_smtp_password" : from_email_smtp_password ,
        }
        

    except:
        # メールの送信文
        message_subject = "REINSスクレイピング定期実行"
        message_body = f"""
            スクレイピングができませんでした。エラーが発生しました。
        """
        file_path = log_txt_path

        # 全てのメールにスクレイピング結果のExcelを送信
        for loop , to_email in enumerate(mail_list):
            cc_mail_row_list = cc_mail_list[loop]
            py_mail.send_py_gmail(
                message_subject , message_body , from_email_smtp_password ,
                from_email , to_email , cc_mail_row_list = cc_mail_row_list ,
                file_path = file_path ,
            )


@app.post("/excel")
def fast_api_excel(api_data_excel: RequestDataExcel):
    log_txt.add_log_txt("2つ目のAPI起動完了")
    print("2つ目のAPI起動完了")
    many_excel_list = api_data_excel.many_excel_list
    search_method_list = api_data_excel.search_method_list
    search_requirement_list = api_data_excel.search_requirement_list

    mail_list = api_data_excel.mail_list
    cc_mail_list = api_data_excel.cc_mail_list
    from_email = api_data_excel.from_email
    from_email_smtp_password = api_data_excel.from_email_smtp_password
    message_subject = "REINSスクレイピング定期実行"
    message_body = ""

    print(f"全てのパラメータを2つ目APIでの取得 が完了")

    try:
        # スクレイピング結果のリストをExcelファイルに保存
        ec.many_list_to_excel(
            many_excel_list , input_reins_excel_path , output_reins_excel_path , search_requirement_list
        )
        print(f"スクレイピング結果のリストをExcelファイルに保存 が完了")

        ##### 最終的にはExcelの定型フォームに貼り付け
        log_txt.add_log_txt("スクレイピング結果をExcelファイルに変更 : 完了")

        search_method_list
        search_solding_list = []
        search_rental_list = []
        for search_method in search_method_list:
            if search_method == "select_solding":
                search_solding_list.append(search_method)
            if search_method == "select_rental":
                search_rental_list.append(search_method)
        
        message_body = f"REINSの定期日時スクレイピング結果のメールです。 \n"
        message_body = message_body + "======================================== \n"
        message_body = message_body + "検索条件 \n"
        
        if search_solding_list != []:
            message_body = message_body + " ①売買検索 : \n"
        for search_method in search_solding_list:
            message_body = search_method + "\n"
        message_body = message_body + "======================================== \n"
        
        if search_rental_list != []:
            message_body = message_body + " ②賃貸検索 : \n"
        for search_requirement in search_rental_list:
            message_body = search_requirement + "\n"
        
        message_body = message_body + "======================================== \n"
        message_body = message_body + "\n"
        message_body = message_body + "※ 検索条件は「01」〜「50」の番号で指定されます \n\n"
        message_body = message_body + "スクレイピング結果は添付のExcelファイルをご覧ください。 \n\n"
        message_body = message_body + "指定日時実行の検索条件を変更する際は、ツール「web_reins」で設定変更が可能です。 \n"

        file_path = output_reins_excel_path
        # 全てのメールにスクレイピング結果のExcelを送信
        for loop , to_email in enumerate(mail_list):
            cc_mail_row_list = cc_mail_list[loop]
            py_mail.send_py_gmail(
                message_subject , message_body , from_email_smtp_password ,
                from_email , to_email , cc_mail_row_list = cc_mail_row_list ,
                file_path = file_path ,
            )

    except:
        print(f"エラー発生")
        log_txt.add_log_txt("エラー発生")
        # メールの送信文
        message_body = f"Excelファイル化ができませんでした。エラーが発生しました。 \n"
        message_body = message_body + "検索条件 \n"
        message_body = message_body + " ①売買検索 : \n"
        for search_method in search_method_list:
            message_body = search_method + "\n"
        message_body = message_body + "\n" + " ②賃貸検索 : \n"
        for search_requirement in search_requirement_list:
            message_body = search_requirement + "\n"
        message_body = message_body + "\n"
        message_body = message_body + "======================================== \n"
        message_body = message_body + "エラーメッセージ : \n"
        message_body = message_body + "======================================== \n\n"
        # message_body = message_body + f"{error_text} \n\n"

        file_path = log_txt_path
        

        # 全てのメールにスクレイピング結果のExcelを送信
        for loop , to_email in enumerate(mail_list):
            cc_mail_row_list = cc_mail_list[loop]
            py_mail.send_py_gmail(
                message_subject , message_body , from_email_smtp_password ,
                from_email , to_email , cc_mail_row_list = cc_mail_row_list ,
                file_path = file_path ,
            )
    print(f"メールの送信が完了")


