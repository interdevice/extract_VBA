"""
Script para extrair código VBA de arquivos Excel (.xlsm, .xls, .xlam)
e salvar em arquivos separados (.bas, .cls, .frm)

INSTRUÇÕES:
1. Coloque seu arquivo Excel (.xlsm, .xls, .xlam) nesta pasta (Documents/excel)
2. Execute: python extrair_vba.py seu_arquivo.xlsm
"""

import os
import sys
from pathlib import Path
import zipfile
import shutil


def extrair_vba_oletools(arquivo_excel, pasta_destino):
    """
    Extrai código VBA usando a biblioteca oletools
    """
    try:
        from oletools.olevba import VBA_Parser
        
        print(f"Processando: {arquivo_excel}")
        
        # Criar pasta de destino se não existir
        Path(pasta_destino).mkdir(parents=True, exist_ok=True)
        
        # Parser VBA
        vba_parser = VBA_Parser(arquivo_excel)
        
        if not vba_parser.detect_vba_macros():
            print("❌ Nenhuma macro VBA encontrada neste arquivo.")
            vba_parser.close()
            return False
        
        print("✅ Macros VBA detectadas!")
        
        # Extrair cada módulo
        modulos_extraidos = 0
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
            if vba_code:
                # Determinar extensão baseada no tipo de módulo
                if vba_filename.startswith("Class"):
                    extensao = ".cls"
                elif vba_filename.startswith("Form") or vba_filename.startswith("UserForm"):
                    extensao = ".frm"
                else:
                    extensao = ".bas"
                
                # Nome do arquivo de saída
                nome_limpo = vba_filename.replace("/", "_").replace("\\", "_")
                arquivo_saida = os.path.join(pasta_destino, f"{nome_limpo}{extensao}")
                
                # Salvar código
                with open(arquivo_saida, "w", encoding="utf-8") as f:
                    f.write(vba_code)
                
                print(f"  📄 Extraído: {nome_limpo}{extensao} ({len(vba_code)} caracteres)")
                modulos_extraidos += 1
        
        vba_parser.close()
        print(f"\n✅ Total: {modulos_extraidos} módulo(s) extraído(s) para: {pasta_destino}")
        return True
        
    except ImportError:
        print("❌ Biblioteca 'oletools' não encontrada.")
        print("   Instale com: pip install oletools")
        return False
    except Exception as e:
        print(f"❌ Erro ao processar arquivo: {e}")
        return False


def extrair_vba_zipfile(arquivo_excel, pasta_destino):
    """
    Método alternativo: extrai VBA de arquivos .xlsm usando zipfile
    (funciona apenas para formato Office Open XML - .xlsm, .xlam)
    """
    try:
        print(f"Tentando método alternativo (zipfile) para: {arquivo_excel}")
        
        # Criar pasta de destino
        Path(pasta_destino).mkdir(parents=True, exist_ok=True)
        
        # Verificar se é um arquivo zip válido
        if not zipfile.is_zipfile(arquivo_excel):
            print("❌ Arquivo não está no formato Office Open XML (.xlsm/.xlam)")
            return False
        
        # Extrair conteúdo VBA
        with zipfile.ZipFile(arquivo_excel, "r") as zip_ref:
            # Procurar por arquivos VBA na pasta xl/vbaProject.bin
            vba_files = [f for f in zip_ref.namelist() if "vbaProject.bin" in f]
            
            if not vba_files:
                print("❌ Nenhum projeto VBA encontrado no arquivo.")
                return False
            
            for vba_file in vba_files:
                # Extrair vbaProject.bin
                vba_bin = os.path.join(pasta_destino, "vbaProject.bin")
                with zip_ref.open(vba_file) as source, open(vba_bin, "wb") as target:
                    target.write(source.read())
                
                print(f"✅ Extraído: vbaProject.bin")
                print(f"   Use oletools para decodificar: olevba {vba_bin}")
                return True
        
        return False
        
    except Exception as e:
        print(f"❌ Erro no método alternativo: {e}")
        return False


def main():
    print("=" * 60)
    print("EXTRATOR DE CÓDIGO VBA DE ARQUIVOS EXCEL")
    print("=" * 60)
    print()
    
    # Verificar argumentos
    if len(sys.argv) < 2:
        print("Uso: python extrair_vba.py <arquivo_excel> [pasta_destino]")
        print()
        print("Exemplo:")
        print("  python extrair_vba.py planilha.xlsm")
        print("  python extrair_vba.py planilha.xlsm ./vba_extracted")
        print()
        print("IMPORTANTE: Coloque o arquivo Excel nesta pasta antes de executar!")
        return
    
    arquivo_excel = sys.argv[1]
    
    # Verificar se arquivo existe
    if not os.path.exists(arquivo_excel):
        print(f"❌ Arquivo não encontrado: {arquivo_excel}")
        print(f"   Certifique-se de que o arquivo está na pasta: {os.getcwd()}")
        return
    
    # Definir pasta de destino
    if len(sys.argv) >= 3:
        pasta_destino = sys.argv[2]
    else:
        # Usar pasta padrão vba_extracted
        nome_base = Path(arquivo_excel).stem
        pasta_destino = f"vba_extracted_{nome_base}"
    
    print(f"Arquivo de entrada: {arquivo_excel}")
    print(f"Pasta de saída: {pasta_destino}")
    print()
    
    # Tentar extrair com oletools (método preferido)
    sucesso = extrair_vba_oletools(arquivo_excel, pasta_destino)
    
    # Se falhar, tentar método alternativo
    if not sucesso:
        print("\nTentando método alternativo...")
        sucesso = extrair_vba_zipfile(arquivo_excel, pasta_destino)
    
    if sucesso:
        print("\n" + "=" * 60)
        print("✅ EXTRAÇÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("❌ Não foi possível extrair o código VBA")
        print("=" * 60)


if __name__ == "__main__":
    main()
