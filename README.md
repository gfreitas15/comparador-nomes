### Comparador de Planilhas

Aplicativo desktop (PyQt5) para comparar dados entre duas planilhas Excel com pré-visualização, normalização e similaridade configurável.

### Requisitos
- Python 3.8+
- Dependências: pandas, PyQt5, rapidfuzz

Instale as dependências:
```bash
pip install -U pandas PyQt5 rapidfuzz
```

### Executar em desenvolvimento
```bash
python comparador.py
```

### Como usar
1. Selecione a PLANILHA 1 e a PLANILHA 2.
2. Marque as colunas de cada planilha que formam a chave.
3. Ajuste similaridade e normalização.
4. Selecione o arquivo de saída (.xlsx).
5. Clique em Comparar. Será mostrada uma prévia de até 20 linhas.
6. Confirme para gerar o Excel final.

Dicas:
- Combine várias colunas quando precisar de contexto (ex.: CPF + NOME).
- Se não conseguir salvar, feche o arquivo de destino no Excel e tente novamente.

### Recursos
- Pré-visualização até 20 linhas, com quebra de texto e colunas autoajustadas.
- Similaridade (0–100) usando RapidFuzz.
- Tema claro/escuro, ajuda e arrastar & soltar arquivos.

### Construir executável (Windows)
Usando o .spec do projeto:
```bash
pyinstaller ComparadorPlanilhas.spec
```

Ou gerando do script:
```bash
pyinstaller --noconfirm --name "ComparadorPlanilhas" --icon "icone.ico" --windowed comparador.py
```

Artefatos em `build/` e executável em `dist/ComparadorPlanilhas.exe`.

### Estrutura
- `comparador.py`: app principal
- `ComparadorPlanilhas.spec`: build PyInstaller
- `icone.ico`: ícone

### Licença
Uso interno/educacional. Ajuste conforme necessário.

