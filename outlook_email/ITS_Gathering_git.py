import win32com.client
import re
import time
from datetime import datetime
import pyodbc

def get_data_from_emails(outlook, target_subject_prefix):
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    results = []
    senders = []
    cc_list = []
    process_codes = []
    original_messages = []

    for message in messages:
        try:
            if message.UnRead and message.Subject.startswith(target_subject_prefix):
                received_time = message.ReceivedTime
                if received_time is None:
                    continue

                print(f"메일 수신 시간 (로컬 시간): {received_time}")
                body = message.Body

                body_match = re.search(r'ProcessCode\s*:\s*(\d+)', body)
                process_code = body_match.group(1) if body_match else "140"

                match = re.search(r'\[Lot\](.*?)\[End\]', body, re.DOTALL)
                if match:
                    content = match.group(1).strip()
                    data = re.findall(r'([\w\-]+)\s*:\s*(\d+)', content)
                    print(data)

                    valid_data = []
                    for key, value in data:
                        if re.match(r'^[A-Za-z0-9\-]+$', key) and re.match(r'^\d+$', value):
                            valid_data.append((key, value))
                        else:
                            print(f"잘못된 포맷 발견: {key} : {value}")
                            valid_data = []
                            break

                    if valid_data:
                        results.append(valid_data)
                        senders.append(message.SenderEmailAddress)
                        cc_list.append(message.CC)
                        process_codes.append(process_code)
                        original_messages.append(message)

                message.UnRead = False

        except Exception as e:
            print(f"이메일 처리 중 오류 발생: {str(e)}")
            message.UnRead = False
            continue

    return results, senders, cc_list, process_codes, original_messages

def query_data_from_sql(data, process_code):
    # SQL Server 연결 정보 설정
    server = ''
    database = ''
    username = ''
    password = ''
    driver = '{ODBC Driver 17 for SQL Server}'

    connection_string = f"DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}"

    # SQL Server 연결
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    print("SQL Server 연결 완료")

    lot_values = [item[0] for item in data]  # key값만 추출

    # 결과 저장을 위한 딕셔너리 (LotNumber를 기준으로 그룹화)
    grouped_result = {}

    # 첫 번째 쿼리 실행
    for lot_value in lot_values:
        lot_value_trimmed = lot_value[:12].strip()  # '-A' 또는 '-00'을 제거하고, 첫 12자리만 추출 후 공백 제거

        query = f"""
            select sh.LotNumber, sh.ProcessCode, p.ProcessName, sh.ComputerName, count(*) as strip_count
            from dbo.pts_StripHistory sh
            left join dbo.Processes p on p.ProcessCode = sh.ProcessCode
            where sh.ProcessCode = {process_code} and sh.LotNumber = substring('{lot_value_trimmed}', 1, 12) + '-00'
            group by sh.LotNumber, sh.ProcessCode, p.ProcessName, sh.ComputerName
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        for row in rows:
            lot_number_trimmed_from_db = row.LotNumber.strip()  # DB에서 가져온 LotNumber에서 공백 제거
            if lot_number_trimmed_from_db not in grouped_result:
                grouped_result[lot_number_trimmed_from_db] = {"lot_info": [], "defect_info": []}
            grouped_result[lot_number_trimmed_from_db]["lot_info"].append(row)

    # 두 번째 쿼리 실행
    for lot_value in lot_values:
        lot_value_trimmed = lot_value[:12].strip()  # '-A' 또는 '-00'을 제거하고, 첫 12자리만 추출 후 공백 제거

        query_2 = f"""
            select substring(s.StripID, 1, 12) as LotNumber,s.DefectCode, d.DefectName, count(*) as defect_count
            from dbo.pts_StripBCL2 s
            left join dbo.Defects d on s.DefectCode = d.DefectCode
            where s.ProcessCode = {process_code}
            and not exists (select 1 from dbo.pts_StripBCL2 t 
                            where t.StripID = s.StripID 
                            and t.PCSCol = s.PCSCol 
                            and t.PCSRow = s.PCSRow 
                            and t.ProcessCode <> s.ProcessCode)
            and s.StripID like substring('{lot_value_trimmed}', 1, 12) + '%'
            group by  substring(s.StripID, 1, 12), s.DefectCode, d.DefectName
            order by s.DefectCode
        """
        cursor.execute(query_2)
        rows = cursor.fetchall()
        if rows:
            for row in rows:
                lot_number_trimmed_from_db = row.LotNumber.strip()  # DB에서 가져온 StripID에서 LotNumber 추출 후 공백 제거
                # 12자리만 비교
                for grouped_lot_number in grouped_result.keys():
                    if lot_number_trimmed_from_db == grouped_lot_number[:12]:
                        grouped_result[grouped_lot_number]["defect_info"].append(row)
        else:
            print(f"LotNumber {lot_value_trimmed}에 대한 Defect 정보가 없습니다.")

    # 커넥션 종료
    cursor.close()
    connection.close()

    return grouped_result


def send_email_with_query_result(original_message, grouped_result, to_email, cc_email=None):
    # Outlook 애플리케이션에 연결
    mail = original_message.Reply()  # 회신 생성

    # 메일 제목 및 수신자 설정
    mail.Subject = f"Re: {original_message.Subject}"
    mail.To = to_email

    # 참조자 설정 (있을 경우)
    if cc_email:
        mail.CC = cc_email

    # SQL 쿼리 결과를 본문에 추가
    body = "\n\n------ 데이터 회신 드립니다. ------\n"
    for lot_number, result in grouped_result.items():
        body += f"{lot_number}\n"

        # Lot 정보 부분
        for row in result["lot_info"]:
            # 'Quad2 474ST' 형식으로 변환
            lot_info = f"{row.ComputerName} {row.strip_count}ST"
            body += f"  {lot_info}\n"

        # Defect 정보 부분
        body += "  Defect 정보:\n"
        for row in result["defect_info"]:
            # '[DefectCode] DefectName : count' 형식으로 변환
            defect_info = f"[{row.DefectCode}] {row.DefectName} : {row.defect_count}"
            body += f"    {defect_info}\n"

        body += "\n" + "-" * 40 + "\n"

    mail.Body = body

    # 메일 전송
    mail.Send()
    print("메일 전송 완료!")


