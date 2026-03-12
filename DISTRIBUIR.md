# 📦 GUIA DE DISTRIBUIÇÃO - Controle DDU

## ✅ OPÇÃO 1: Distribuir com Python Instalado (MAIS RÁPIDO AGORA)

**Recomendado enquanto o executável está sendo compilado**

### Passo 1: No computador de origem
Copie toda a pasta:
```
C:\Users\Lucas\Desktop\Projetos\Python\Controle_DDU
```

### Passo 2: No computador de destino
Copie a pasta para a localização desejada e execute:

```bash
# Instale as dependências
pip install -r requirements.txt

# Execute o programa
python app.py
```

**Requisitos no computador de destino:**
- Windows 7+ ou qualquer SO com Python
- Python 3.10+ (baixar em python.org)

---

## ⏳ OPÇÃO 2: Executável Portável (EM PREPARAÇÃO)

O arquivo `Controle_DDU.exe` está sendo gerado e aparecerá em:
```
dist\Controle_DDU.exe
```

Quando pronto, esse arquivo:
- ✅ Não precisa de Python
- ✅ Funciona em qualquer Windows
- ✅ Pode ser copiado isoladamente
- ✅ Tudo incluído em um arquivo

---

## 📋 O QUE DISTRIBUIR

### Opção 1 (Com Python):
```
Controle_DDU/
├── app.py
├── requirements.txt
├── documents.db (será criado automaticamente)
├── pmmg.png
└── ... outros arquivos
```

### Opção 2 (SEM Python - quando pronto):
```
Controle_DDU.exe  ← Único arquivo necessário
```

---

## 🚀 INSTALAÇÃO NO OUTRO COMPUTADOR

### Com Python (Fácil):
1. Copie a pasta inteira
2. Abra CMD/PowerShell no local da pasta
3. Execute:
   ```bash
   pip install -r requirements.txt
   python app.py
   ```

### Sem Python (Quando .exe estiver pronto):
1. Pegue o arquivo `Controle_DDU.exe`
2. Copie para qualquer lugar
3. Clique 2x para executar

---

## ❓ DÚVIDAS FREQUENTES

**P: Python não está instalado no outro computador?**  
R: Use o `Controle_DDU.exe` (quando disponível) - não precisa de nada extra.

**P: Onde fica guardado o banco de dados?**  
R: Na mesma pasta do programa (`documents.db`). Não se perde ao desinstalar.

**P: Posso usar em rede/servidor?**  
R: Sim, coloque a pasta em local compartilhado. Todos podem acessar o mesmo banco de dados.

**P: Preciso fazer backup?**  
R: Sim! Copie o arquivo `documents.db` regularmente.

---

## 🔧 STATUS DA COMPILAÇÃO

**Executável em processo de compilação...**
- ⏳ Tempo estimado: 5-10 minutos
- 📍 Localização final: `dist\Controle_DDU.exe`
- ✓ Quando pronto, aparecerá na pasta `dist/`

---

**Desenvolvido por:** Lucas Pires Franco - Engenheiro de Computação
