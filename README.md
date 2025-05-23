# 📄 Editor de Documentos Word

E aí pessoal! 👋

Criei esse codigo porque tava cansado de ficar atualizando manualmente um monte de documentos Word na minha empresa. Todo mês era a mesma coisa: abrir arquivo por arquivo, mudar os textos nos rodapés e depois converter tudo pra PDF um saco

Aí um dia resolvi automatizar isso tudo com Python. O que começou como um script simples acabou virando essa bomba com interface bonitinha e tudo. Como me ajudou muito no dia a dia, resolvi compartilhar - vai que ajuda mais alguém com o mesmo problema, né? 

## ✨ O que ele faz?

- 🔄 Atualiza vários documentos Word de uma vez
- 📝 Mexe principalmente nos rodapés (que era meu maior problema)
- 🔍 Procura e substitui textos específicos
- 📁 Processa a pasta inteira de uma vez
- 🎯 Interface simples de usar
- 📊 Mostra tudo que está acontecendo
- 🚀 Já converte pra PDF no final

## 🛠️ Tecnologias 

- Python 3.x
- CustomTkinter (pra interface ficar bonita)
- python-docx (pra mexer nos arquivos do Word)
- pywin32 (pra automatizar o Word)

## 📋 Pra funcionar você precisa ter:

- Python 3.x instalado
- Microsoft Word instalado
- Feito para Windows

## ⚙️ Como usar

### Se você manja de Python:

1. Baixa os arquivos:
```bash
git clone https://github.com/seu-usuario/editor-documentos-word.git
```

2. Instala as dependências:
```bash
pip install -r requirements.txt
```

3. Roda o programa:
```bash
python word_processor.py
```

### Se você quer só usar:

1. Baixa a versão compilada (executável) nas releases
2. Extrai o arquivo
3. Executa o `word_processor.exe`

### Se você quer compilar:

Já deixei o arquivo spec pronto! É só rodar:
```bash
pyinstaller word_processor.spec
```
O executável vai estar na pasta `dist`.

### Como funciona:

1. Na tela que abrir:
   - Coloca os textos que quer trocar
   - Escolhe a pasta com os arquivos Word
   - Escolhe onde salvar os PDFs
   - Clica em "Iniciar Processamento"
   - Pronto! 🎉

## 💡 Por que usar?

Cara, isso me economizou MUITAS horas de trabalho chato e repetitivo. Se você também precisa:
- ⏱️ Atualizar vários documentos de uma vez
- 🎯 Evitar erros de digitação
- 📊 Ter certeza que nada foi esquecido
- 🔄 Fazer isso todo mês/semana


## 📝 Licença

Projeto sob licença MIT - pode usar à vontade!

## 🙋‍♂️ Sobre

Fiz esse código pra resolver um problema real na minha empresa e resolvi compartilhar. Se te ajudou, deixa uma ⭐ no projeto! 