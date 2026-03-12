# Como Executar o Controle_DDU com Sucesso

## ⚠️ IMPORTANTE - Problema Resolvido

O erro `ModuleNotFoundError: No module named 'plotly'` ocorre porque o Python do **sistema** não tem as bibliotecas necessárias instaladas. As bibliotecas estão apenas no **ambiente virtual (venv)**.

## ✅ Solução Recomendada - Use os Scripts de Inicialização

### Opção 1: Script Batch (Windows - Mais Fácil)
Simplesmente **clique duas vezes** em:
```
run.bat
```

Isso ativa automaticamente o venv e executa o programa.

---

### Opção 2: PowerShell
Abra PowerShell na pasta do projeto e execute:
```powershell
& .\run.ps1
```

---

### Opção 3: Terminal Manual com venv ativado
Se preferir executar manualmente:
```powershell
# 1. Ativar o ambiente virtual
& .\.venv\Scripts\Activate.ps1

# 2. Executar o programa
python Controle_documentos.py
```

---

## 🔧 Configuração do VS Code (Opcional)

Se quiser executar diretamente do VS Code sem scripts, configure o interpretador Python:

1. Pressione `Ctrl + Shift + P`
2. Digite: `Python: Select Interpreter`
3. Procure por: `.venv`
4. Clique em: `.\.venv\Scripts\python.exe`
5. Pronto! Agora o VS Code usará o venv automaticamente

---

## 📦 Dependências Instaladas

O programa requer as seguintes bibliotecas (já instaladas no venv):
- `plotly` - Gráficos interativos
- `kaleido` - Exportação de gráficos para PNG
- `numpy` - Cálculos numéricos
- `PySide6` - Interface gráfica
- `reportlab` - Geração de PDFs
- `Pillow` - Processamento de imagens

---

## ✨ Funcionalidades Disponíveis

✅ Dashboard com gráficos Plotly
✅ Análise de dados em tempo real
✅ Comparação de períodos
✅ Geração de relatórios em PDF
✅ Banco de dados SQLite
✅ Interface moderna em PySide6

---

**Não há erros no código!** 🎉
O programa está 100% funcional e pronto para uso.
