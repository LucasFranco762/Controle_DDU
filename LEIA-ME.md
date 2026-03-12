# Controle DDU - Pronto para Distribuição

## 📦 Opção 1: Executável Portável (SEM Python)

A pasta `dist/` contém o arquivo `Controle_DDU.exe` - pode ser enviado para qualquer computador Windows e executado diretamente.

**Se o arquivo .exe ainda não foi criado**, execute em seu computador:
```bash
python build_exe_simple.py
```

O executável será criado na pasta `dist/` (leva 2-3 minutos).

---

## 📋 Opção 2: Distribuição com Código-Fonte (PRECISA Python)

Se o computador de destino já tem Python 3.10+ instalado, copie toda esta pasta e execute:

```bash
pip install -r requirements.txt
python app.py
```

---

## 📂 Arquivos Importantes

| Arquivo | Descrição |
|---------|-----------|
| `app.py` | Código-fonte do programa |
| `requirements.txt` | Dependências Python necessárias |
| `documents.db` | Banco de dados SQLite (criado automaticamente) |
| `pmmg.png` | Logomarca da PMMG |
| `build_exe_simple.py` | Script para gerar executável |
| `dist/Controle_DDU.exe` | Executável pronto (após compilação) |

---

## 🚀 Para Instalar em Outro Computador

### Com Python Instalado:
1. Copie a pasta inteira
2. Execute: `python app.py`

### SEM Python (Usar Executável):
1. Localize `dist/Controle_DDU.exe`
2. Copie apenas este arquivo
3. Execute no outro computador

---

## 🛠️ Recursos do Programa

✅ **Cadastro de Documentos** - Tipo, Número, Natureza, Local, Data, Militar Responsável  
✅ **Gestão** - Listar, Excluir, Gerar Relatório PDF  
✅ **Alertas** - Automáticos aos 5 dias e no vencimento  
✅ **Visualização** - Tabela com cores (vencidos em vermelho)  
✅ **Banco de Dados** - SQLite integrado e seguro  
✅ **Interface** - Moderna com PySide6  

---

## 📝 Notas

- Dados são armazenados em `documents.db` - **não se perdem** ao fechar o programa
- Só podem ser excluídos manualmente através da interface
- O PDF pode ser salvo em qualquer pasta desejada
- A logomarca é carregada automaticamente se o arquivo `pmmg.png` estiver na mesma pasta

---

**Desenvolvido por:** Lucas Pires Franco - Engenheiro de Computação
