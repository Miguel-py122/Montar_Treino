# Planner de Treinos

Aplicacao web front-end para montar fichas de treino com preenchimento rapido, busca de exercicios por agrupamento muscular, salvamento automatico no navegador e exportacao profissional.

## Recursos

- Tabela interativa com edicao direta
- Treinos separados em abas
- Biblioteca ampliada de exercicios por agrupamento muscular
- Busca pesquisavel com sugestoes automaticas por agrupamento e troca rapida de exercicio
- Dropdown visual customizado para agrupamento muscular
- Presets de divisao do dia e mapa corporal interativo
- Mapa corporal em SVG com alternancia frente e costas
- Presets com exercicios-base sugeridos automaticamente
- Duplicacao rapida de linhas de exercicio
- Salvamento automatico com LocalStorage
- Exportacao principal em PDF do treino ativo ou de todos os treinos
- Remocao do treino atual com confirmacao
- Layout responsivo

## Como usar

1. Abra o arquivo `index.html` no navegador.
2. Use os presets de divisao do dia ou o mapa corporal para acelerar a montagem.
3. Escolha o agrupamento muscular no dropdown visual e pesquise os exercicios da linha.
4. Use o botao "Duplicar" para repetir uma linha preenchida em um clique.
5. Preencha os campos da linha como em uma planilha.
6. Use "Baixar PDF" para exportar o treino ativo ou "Baixar tudo" para gerar um PDF consolidado.
7. Use "Remover treino" quando quiser apagar a ficha atual com confirmacao.

## Estrutura

- `index.html`: estrutura da aplicacao
- `styles.css`: interface e responsividade
- `script.js`: logica, biblioteca de exercicios, armazenamento local, importacao e exportacao
- `exemplos/`: arquivos-modelo para importacao