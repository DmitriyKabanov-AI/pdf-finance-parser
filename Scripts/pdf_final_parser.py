import re
import pdfplumber
from pathlib import Path
import pandas as pd

def extract_transactions(pdf_path):
    transactions = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                # Начало транзакции: строка начинается с даты и содержит знак рубля
                if re.match(r'^\d{2}\.\d{2}\.\d{4}', line) and '₽' in line:
                    block = [line]
                    j = i + 1
                    # Собираем следующие строки, пока не встретим новую дату
                    while j < len(lines) and not re.match(r'^\d{2}\.\d{2}\.\d{4}', lines[j].strip()):
                        block.append(lines[j].strip())
                        j += 1
                    full = ' '.join(block)
                    
                    # Парсим объединённую строку
                    # Формат: дата1 дата2 сумма1 ₽ сумма2 ₽ описание ... (время может быть разбито)
                    # Пример: "29.12.2025 29.12.2025 -31 571.00 ₽ -31 571.00 ₽ Внутренний перевод на 0389 11:22 11:23 договор 0355576830"
                    # Время может быть как после дат, так и в конце.
                    # Ищем шаблон: дата пробел дата пробел число ₽ пробел число ₽ ... и в конце возможно время время
                    # Упростим: извлечём даты, суммы, а всё остальное — описание, в котором могут быть времена и номер карты.
                    # Лучше использовать два этапа: сначала выделить даты и суммы, потом остаток.
                    
                    # Ищем даты (две подряд)
                    date_pattern = r'(\d{2}\.\d{2}\.\d{4})\s+(\d{2}\.\d{2}\.\d{4})'
                    date_match = re.search(date_pattern, full)
                    if not date_match:
                        i = j
                        continue
                    date_op, date_wr = date_match.groups()
                    
                    # Ищем суммы (две подряд с ₽)
                    sum_pattern = r'([+-]?[\d\s.,]+)₽\s+([+-]?[\d\s.,]+)₽'
                    sum_match = re.search(sum_pattern, full)
                    if not sum_match:
                        i = j
                        continue
                    sum_op_raw, sum_card_raw = sum_match.groups()
                    sum_op = sum_op_raw.replace(' ', '').replace(',', '.')
                    sum_card = sum_card_raw.replace(' ', '').replace(',', '.')
                    
                    # Остаток строки после сумм — это описание + возможно время + номер карты
                    after_sums = full[sum_match.end():].strip()
                    # Из описания извлекаем время (два раза по ЧЧ:ММ) и номер карты (4 цифры в конце)
                    time_pattern = r'(\d{2}:\d{2})\s+(\d{2}:\d{2})'
                    time_match = re.search(time_pattern, after_sums)
                    if time_match:
                        time_op, time_wr = time_match.groups()
                        # Удаляем время из описания
                        after_sums = after_sums[:time_match.start()] + after_sums[time_match.end():]
                    else:
                        # Если время не найдено, возможно оно в начале описания
                        time_op = time_wr = None
                    
                    # Номер карты (4 цифры) в конце описания
                    card_match = re.search(r'(\d{4})$', after_sums)
                    card = card_match.group(1) if card_match else None
                    if card:
                        after_sums = after_sums[:card_match.start()].strip()
                    
                    description = after_sums.strip()
                    
                    # Собираем транзакцию
                    trans = {
                        'datetime_operation': f"{date_op} {time_op}" if time_op else date_op,
                        'datetime_writeoff': f"{date_wr} {time_wr}" if time_wr else date_wr,
                        'amount_operation': float(sum_op) if sum_op else None,
                        'amount_card': float(sum_card) if sum_card else None,
                        'description': description,
                        'card_number': card
                    }
                    transactions.append(trans)
                    i = j
                else:
                    i += 1
    return pd.DataFrame(transactions)

def main():
    # Ищем папку Statements рядом со скриптом
    script_dir = Path(__file__).parent.parent  # поднимаемся из Scripts/ в Project1/
    folder = script_dir / "Statements"
    
    if not folder.exists():
        print(f"Папка не найдена: {folder}")
        return
    
    pdfs = list(folder.glob('*.pdf')) + list(folder.glob('*.PDF'))
    pdfs = sorted(set(pdfs))
    if not pdfs:
        print("Нет PDF.")
        return
    
    all_dfs = []
    for pdf in pdfs:
        print(f"Обработка {pdf.name}...")
        df = extract_transactions(pdf)
        if not df.empty:
            df.insert(0, 'source_file', pdf.name)
            all_dfs.append(df)
            print(f"  Найдено {len(df)} транзакций")
        else:
            print("  Транзакций не найдено")
    
    if not all_dfs:
        print("Ничего не найдено.")
        return
    
    final_df = pd.concat(all_dfs, ignore_index=True)
    output = folder / "transactions_consolidated.xlsx"
    final_df.to_excel(output, index=False)
    print(f"\nСохранено {len(final_df)} транзакций в {output}")

if __name__ == "__main__":
    main()