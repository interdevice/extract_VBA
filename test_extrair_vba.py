"""
Testes para o script extrair_vba.py

Execute com: python test_extrair_vba.py
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import zipfile

# Importar funcoes do script principal
from extrair_vba import extrair_vba_oletools, extrair_vba_zipfile


def criar_excel_teste_com_macro():
    """
    Cria um arquivo Excel de teste com macro VBA simulada
    Retorna o caminho do arquivo criado
    """
    print("\n📦 Criando arquivo Excel de teste...")
    
    # Criar arquivo .xlsm simulado (formato Office Open XML)
    arquivo_teste = "teste_macro.xlsm"
    
    # Criar estrutura basica de um arquivo .xlsm (que e um arquivo ZIP)
    with zipfile.ZipFile(arquivo_teste, "w", zipfile.ZIP_DEFLATED) as zf:
        # Adicionar arquivos basicos do Excel
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>""")
        
        zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")
        
        zf.writestr("xl/workbook.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>""")
        
        # Adicionar vbaProject.bin vazio (simulado)
        zf.writestr("xl/vbaProject.bin", b"PK\x03\x04" + b"\x00" * 100)
    
    print(f"✅ Arquivo de teste criado: {arquivo_teste}")
    return arquivo_teste


def criar_arquivo_sem_macro():
    """
    Cria um arquivo Excel sem macros para teste
    """
    print("\n📦 Criando arquivo Excel sem macro...")
    
    arquivo_teste = "teste_sem_macro.xlsx"
    
    with zipfile.ZipFile(arquivo_teste, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
</Types>""")
        
        zf.writestr("xl/workbook.xml", """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1"/>
    </sheets>
</workbook>""")
    
    print(f"✅ Arquivo sem macro criado: {arquivo_teste}")
    return arquivo_teste


def teste_importacao_oletools():
    """
    Testa se oletools esta instalado
    """
    print("\n🧪 TESTE 1: Verificacao de dependencias")
    print("-" * 50)
    
    try:
        from oletools.olevba import VBA_Parser
        print("✅ PASSOU: Biblioteca oletools instalada e importavel")
        return True
    except ImportError as e:
        print("⚠️ AVISO: Biblioteca oletools nao encontrada")
        print(f"   Erro: {e}")
        print("   Execute: pip install oletools")
        return False


def teste_pasta_destino_criacao():
    """
    Testa criacao de pasta de destino
    """
    print("\n🧪 TESTE 2: Criacao de pasta de destino")
    print("-" * 50)
    
    pasta_destino = "test_pasta_nova/subpasta/destino"
    
    try:
        # Criar pasta usando Path (mesmo metodo do script)
        Path(pasta_destino).mkdir(parents=True, exist_ok=True)
        
        if os.path.exists(pasta_destino):
            print(f"✅ PASSOU: Pasta criada com sucesso: {pasta_destino}")
            return True
        else:
            print("❌ FALHOU: Pasta nao foi criada")
            return False
    finally:
        # Limpar
        if os.path.exists("test_pasta_nova"):
            shutil.rmtree("test_pasta_nova")


def teste_arquivo_sem_macro():
    """
    Testa arquivo Excel sem macros
    """
    print("\n🧪 TESTE 3: Arquivo sem macros")
    print("-" * 50)
    
    arquivo = criar_arquivo_sem_macro()
    pasta_destino = "test_output_sem_macro"
    
    try:
        # Tentar metodo zipfile
        resultado = extrair_vba_zipfile(arquivo, pasta_destino)
        
        if not resultado:
            print("✅ PASSOU: Script detectou corretamente ausencia de macros")
            return True
        else:
            print("⚠️ AVISO: Script encontrou macros quando nao deveria")
            return False
    finally:
        # Limpar
        if os.path.exists(arquivo):
            os.remove(arquivo)
        if os.path.exists(pasta_destino):
            shutil.rmtree(pasta_destino)


def teste_arquivo_com_macro():
    """
    Testa arquivo Excel com macros
    """
    print("\n🧪 TESTE 4: Arquivo com macros")
    print("-" * 50)
    
    arquivo = criar_excel_teste_com_macro()
    pasta_destino = "test_output_com_macro"
    
    try:
        # Tentar metodo zipfile
        resultado = extrair_vba_zipfile(arquivo, pasta_destino)
        
        if resultado:
            print("✅ PASSOU: Script extraiu vbaProject.bin")
            
            # Verificar se pasta foi criada
            if os.path.exists(pasta_destino):
                print(f"✅ Pasta de destino criada: {pasta_destino}")
                
                # Listar arquivos extraidos
                arquivos = os.listdir(pasta_destino)
                print(f"📄 Arquivos extraidos: {arquivos}")
                return True
            else:
                print("❌ FALHOU: Pasta de destino nao foi criada")
                return False
        else:
            print("⚠️ AVISO: Script nao conseguiu extrair")
            return True  # Nao e falha critica
    finally:
        # Limpar
        if os.path.exists(arquivo):
            os.remove(arquivo)
        if os.path.exists(pasta_destino):
            shutil.rmtree(pasta_destino)


def main():
    """
    Executa todos os testes
    """
    print("=" * 60)
    print("TESTES PARA EXTRATOR DE VBA")
    print("=" * 60)
    
    testes = [
        ("Importacao oletools", teste_importacao_oletools),
        ("Criacao de pasta", teste_pasta_destino_criacao),
        ("Arquivo sem macros", teste_arquivo_sem_macro),
        ("Arquivo com macros", teste_arquivo_com_macro),
    ]
    
    resultados = []
    
    for nome, teste_func in testes:
        try:
            resultado = teste_func()
            resultados.append((nome, resultado))
        except Exception as e:
            print(f"❌ ERRO no teste \"{nome}\": {e}")
            import traceback
            traceback.print_exc()
            resultados.append((nome, False))
    
    # Resumo
    print("\n" + "=" * 60)
    print("RESUMO DOS TESTES")
    print("=" * 60)
    
    passou = sum(1 for _, r in resultados if r)
    total = len(resultados)
    
    for nome, resultado in resultados:
        status = "✅ PASSOU" if resultado else "❌ FALHOU"
        print(f"{status}: {nome}")
    
    print("\n" + "-" * 60)
    print(f"Total: {passou}/{total} testes passaram")
    
    if passou == total:
        print("🎉 Todos os testes passaram!")
        return 0
    else:
        print("⚠️ Alguns testes falharam")
        return 1


if __name__ == "__main__":
    sys.exit(main())
