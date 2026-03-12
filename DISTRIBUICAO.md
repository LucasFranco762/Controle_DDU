# Controle DDU - Distribuição

## Para Distribuir o Programa

### Opção 1: Arquivo Executável (Recomendado)
O arquivo `Controle_DDU.exe` pode ser executado em qualquer computador Windows **sem precisar instalar Python**.

**Passos:**
1. Copie o arquivo `Controle_DDU.exe` da pasta `dist/`
2. Envie para outro computador
3. Execute o arquivo - pronto!

### Opção 2: Com Código-Fonte (Precisa Python)
Se quiser distribuir com o código-fonte:

**Requisitos no outro computador:**
- Python 3.10+ instalado
- PySide6: `pip install PySide6`
- reportlab: `pip install reportlab`

**Passos:**
1. Copie a pasta inteira do projeto
2. Execute: `python app.py`

## Criando o Executável

Se precisar recriar o executável, execute:

```bash
C:/Users/Lucas/Desktop/Projetos/Python/Controle_DDU/.venv/Scripts/pyinstaller --name Controle_DDU --onefile --windowed --add-data "C:\Users\Lucas\Desktop\Projetos\Python\Controle_DDU/documents.db:." --add-data "C:\Users\Lucas\Desktop\Projetos\Python\Controle_DDU/pmmg.png:." --distpath "C:\Users\Lucas\Desktop\Projetos\Python\Controle_DDU\dist" "C:\Users\Lucas\Desktop\Projetos\Python\Controle_DDU\app.py"
```

## Estrutura de Arquivos

```
Controle_DDU/
├── app.py                    # Código-fonte
├── requirements.txt          # Dependências
├── documents.db              # Banco de dados (criado automaticamente)
├── pmmg.png                  # Logomarca
├── Controle_DDU.spec         # Configuração PyInstaller
└── dist/
    └── Controle_DDU.exe      # Executável pronto para distribuir
```

## Recursos do Programa

✅ Cadastrar novos documentos  
✅ Excluir documentos  
✅ Listar documentos (separados: a vencer e vencidos)  
✅ Gerar relatório em PDF  
✅ Alertas automáticos (5 dias antes e no vencimento)  
✅ Banco de dados SQLite integrado  
✅ Interface moderna com PySide6  

## Suporte

Desenvolvido por: Lucas Pires Franco - Engenheiro de Computação
