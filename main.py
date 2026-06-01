import os
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS

# ==========================================
# CONFIGURAÇÃO DE CAMINHOS (DINÂMICOS)
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_FOTOS = os.path.join(BASE_DIR, "fotos_para_verificar")
CAMINHO_SAIDA_EXCEL = os.path.join(BASE_DIR, "auditoria_metadados.xlsx")

def extrair_detalhes_reais(img):
    """
    Varre a imagem à procura de metadados críticos (Make, Model, DateTime).
    Retorna uma lista de dados encontrados para auditoria.
    """
    detalhes = []
    
    # 1. Extração de dados EXIF (Padrão para JPG/TIFF)
    try:
        exif_data = img.getexif()
        if exif_data:
            for tag_id, valor in exif_data.items():
                tag_nome = TAGS.get(tag_id, tag_id)
                # Filtro de tags relevantes para comprovação de autoria
                if tag_nome in ['Make', 'Model', 'DateTime', 'Software'] and valor:
                    valor_limpo = str(valor).strip()
                    if valor_limpo and not valor_limpo.isspace():
                        detalhes.append(f"{tag_nome}: {valor_limpo}")
    except Exception:
        pass
        
    # 2. Extração de metadados internos (PNG, WebP ou Softwares de Edição)
    try:
        if hasattr(img, 'info') and img.info:
            for k, v in img.info.items():
                if k in ['comment', 'Description', 'Software', 'Creation Time'] and v:
                    v_limpo = str(v).strip()
                    if v_limpo and not v_limpo.isspace():
                        detalhes.append(f"{k}: {v_limpo}")
    except Exception:
        pass
        
    return detalhes

def realizar_auditoria():
    if not os.path.exists(PASTA_FOTOS):
        os.makedirs(PASTA_FOTOS)
        print(f"Pasta criada. Coloca as fotos em: {PASTA_FOTOS}")
        return

    extensoes_suportadas = ('.jpg', '.jpeg', '.png', '.tiff', '.webp')
    dados_auditoria = []
    
    print(f"A iniciar auditoria de metadados em: {PASTA_FOTOS}...\n")
    
    arquivos = [f for f in os.listdir(PASTA_FOTOS) if f.lower().endswith(extensoes_suportadas)]
    
    for arquivo in arquivos:
        caminho_completo = os.path.join(PASTA_FOTOS, arquivo)
        nome_sem_extensao, _ = os.path.splitext(arquivo)
        
        try:
            with Image.open(caminho_completo) as img:
                detalhes = extrair_detalhes_reais(img)
                
                tem_metadados = "Sim" if detalhes else "Não"
                info_metadados = ", ".join(detalhes)[:200] # Limite de caracteres para o Excel
                
        except Exception as e:
            tem_metadados = "Erro"
            info_metadados = f"Erro de Leitura: {str(e)}"
        
        dados_auditoria.append({
            "Código da Foto": nome_sem_extensao,
            "Contém Metadados": tem_metadados,
            "Detalhes Encontrados": info_metadados
        })
        print(f" -> {nome_sem_extensao} | Metadados: {tem_metadados}")

    # Gerar Relatório Excel
    if dados_auditoria:
        df = pd.DataFrame(dados_auditoria)
        with pd.ExcelWriter(CAMINHO_SAIDA_EXCEL, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Auditoria EXIF")
            
            # Ajuste automático de colunas
            worksheet = writer.sheets["Auditoria EXIF"]
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                worksheet.column_dimensions[col[0].column_letter].width = max_len + 5

        print(f"\n[SUCESSO] Relatório gerado: {CAMINHO_SAIDA_EXCEL}")

if __name__ == '__main__':
    realizar_auditoria()
