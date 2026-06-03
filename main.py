import subprocess

print("Запуск парсинга PDF...")
subprocess.run(["python", "Scripts/pdf_final_parser.py"])

print("\nЗапуск генерации дашборда...")
subprocess.run(["python", "Scripts/generate_html_report.py"])

print("\nГотово!")