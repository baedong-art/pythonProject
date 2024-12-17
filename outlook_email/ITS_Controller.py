import time
import win32com.client
from ITS_Gathering import get_data_from_emails, query_data_from_sql, send_email_with_query_result
from ITS_Upload import load_config, process_emails
from ITS_Download import get_data_from_emails as get_download_data, query_data_from_sql as query_download_data, save_results_to_memory, send_email_with_attachment

# Outlook 초기화
outlook = win32com.client.Dispatch("Outlook.Application")

print("ITS_Gathering 시작..")

def run_periodically(outlook, interval_seconds=60):
    while True:
        print("이메일 확인 중...")

        try:
            unread_messages = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items.Restrict("[Unread]=True")
            unread_count = unread_messages.Count
            print(f"가져온 읽지 않은 메일 수: {unread_count}")

            # for문을 돌리기 위해서 읽어온 메일을 리스트에 저장 ** 리스트 형태로 저장하지 않으면 for문이 한번만돌고 끝남
            unread_messages = list(unread_messages)


            for message in unread_messages:
                if "[ITS_Upload]" in (message.Subject or ""):
                    shared_folder_path = load_config()
                    if shared_folder_path:
                        process_emails(outlook, shared_folder_path)
                elif "[ITS_Gathering]" in (message.Subject or ""):
                    results, senders, cc_list, process_codes, original_messages = get_data_from_emails(outlook, "[ITS_Gathering]")
                    for result, sender_email, cc_email, process_code, original_message in zip(results, senders, cc_list, process_codes, original_messages):
                        if not result or not sender_email or not process_code:
                            print("데이터, 발신자 이메일 또는 ProcessCode가 없습니다. 건너뜁니다.")
                            continue

                        grouped_result = query_data_from_sql(result, process_code)
                        send_email_with_query_result(original_message, grouped_result, sender_email, cc_email)


                elif "[ITS_Download]" in (message.Subject or ""):

                    results, senders, cc_list, original_messages = get_download_data(outlook,
                                                                                                    "[ITS_Download]")
                    for result, sender_email, cc_email, original_message in zip(results, senders, cc_list,original_messages):
                        if not result or not sender_email:
                            print("데이터, 발신자 이메일  없습니다. 건너뜁니다.")
                            continue

                        grouped_result = query_download_data(result)

                        file_path = save_results_to_memory(grouped_result)
                        send_email_with_attachment(original_message, file_path, sender_email, cc_email)
        except Exception as e:
            print(f"전체 처리 중 오류 발생: {str(e)}")

        print(f"{interval_seconds}초 후에 다시 확인합니다...")
        time.sleep(interval_seconds)

if __name__ == "__main__":
    run_periodically(outlook, interval_seconds=60)