import os
import json

def load_config():
    try:
        with open("config.json", "r", encoding="utf-8") as config_file:
            config_data = json.load(config_file)
            return config_data["shared_folder_path"]
    except Exception as e:
        print(f"Config 파일 읽기 중 오류 발생: {e}")
        return None

def process_emails(outlook, shared_folder_path):
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    unread_messages = inbox.Items.Restrict("[Unread]=True")

    for message in unread_messages:
        try:
            if "[ITS_Upload]" in (message.Subject or ""):
                print(f"처리 중인 이메일: {message.Subject}")
                attachments = message.Attachments
                save_success = True

                for attachment in attachments:
                    property_accessor = attachment.PropertyAccessor
                    content_id = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")

                    if not content_id:
                        save_path = os.path.join(shared_folder_path, attachment.FileName)
                        try:
                            attachment.SaveAsFile(save_path)
                            print(f"첨부파일 저장 완료: {save_path}")
                        except Exception as e:
                            print(f"첨부파일 저장 실패: {e}")
                            save_success = False

                reply = message.Reply()
                if save_success:
                    reply.HTMLBody = f"<p>안녕하세요,</p><p>첨부파일이 성공적으로 공유 폴더에 업로드되었습니다.</p><p>감사합니다.</p>" + reply.HTMLBody
                    print("성공 회신 메일 작성 완료.")
                else:
                    reply.HTMLBody = f"<p>안녕하세요,</p><p>첨부파일을 공유 폴더에 업로드하지 못했습니다. 다시 확인 부탁드립니다.</p><p>감사합니다.</p>" + reply.HTMLBody
                    print("실패 회신 메일 작성 완료.")

                if message.CC:
                    reply.CC = message.CC

                reply.Send()
                print("회신 메일 전송 완료.")
                message.Unread = False

        except Exception as e:
            print(f"이메일 처리 중 오류 발생: {e}")