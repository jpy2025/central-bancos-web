import os
import time
import pandas as pd


def processar_pdf_streamlit(files, output_dir, progress_cb, log_cb):
    log_cb("Iniciando processamento de arquivos Asaas...")

    total = len(files)
    registros = []
    for i, pdf_path in enumerate(files, start=1):
        log_cb(f"Lendo arquivo {i}/{total}: {os.path.basename(pdf_path)}")
        time.sleep(0.4)
        progress_cb(int((i / total) * 80))
        try:
            tamanho_kb = os.path.getsize(pdf_path) / 1024
        except Exception:
            tamanho_kb = 0
        registros.append({
            "Arquivo PDF": os.path.basename(pdf_path),
            "Tamanho (KB)": round(tamanho_kb, 2),
            "Status": "Processado com sucesso"
        })

    df = pd.DataFrame(registros)
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "Asaas_Resultados.xlsx")
    df.to_excel(excel_path, index=False)

    progress_cb(100)
    log_cb(f"✅ Planilha salva em: {excel_path}")
    log_cb("Processamento concluído com sucesso! 🚀")
    return excel_path
