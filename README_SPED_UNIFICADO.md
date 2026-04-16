# SPED Unificado

## O que este projeto faz
- Converte um arquivo `SPED.txt` para Excel.
- Processa automaticamente as abas `D100` e `C100`.
- Permite uso por interface com Streamlit.
- Pode ser iniciado com duplo clique pelo arquivo `Iniciar_SPED_Streamlit.bat`.

## Arquivos
- `sped_unificado_streamlit.py` → aplicação principal
- `Iniciar_SPED_Streamlit.bat` → inicializador com duplo clique no Windows
- `requirements_sped.txt` → dependências

## Como instalar
No Prompt de Comando, dentro da pasta dos arquivos:

```bat
pip install -r requirements_sped.txt
```

## Como abrir
### Opção 1: duplo clique
Dê duplo clique em `Iniciar_SPED_Streamlit.bat`.

### Opção 2: manual
```bat
streamlit run sped_unificado_streamlit.py
```

## Melhorias aplicadas
- remove caminho fixo do TXT
- elimina perguntas no terminal
- unifica conversão e processamento
- permite baixar o arquivo pronto pela tela
- permite salvar automaticamente em uma pasta local
- separa regras de negócio da interface

## Sugestão futura
Se quiser, o próximo passo é transformar isso em `.exe` para abrir sem precisar do terminal.
