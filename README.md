Introdução
Este projeto de previsão do tempo é uma aplicação gráfica desenvolvida em Python com Tkinter, que permite ao usuário consultar rapidamente as condições adversas de uma cidade específica. Utilizando a API do OpenWeatherMap, a aplicação obtém dados sobre a temperatura e a umidade da cidade informada pelo usuário. A previsão é exibida em uma janela do navegador, que fecha automaticamente após alguns segundos, e os dados são salvos em uma planilha para registro e consulta futura.

Objetivo
O principal objetivo do projeto é fornecer uma interface intuitiva e eficiente para o usuário observar o tempo. O sistema também armazena os dados coletados em uma planilha Excel, facilitando a análise histórica das condições climáticas registradas.

Tecnologias Utilizadas
Python : Linguagem principal para o desenvolvimento da aplicação.
Tkinter : Biblioteca para interface gráfica, usada para criar janelas e botões de interação.
OpenWeatherMap API : API de previsão do tempo, que fornece dados meteorológicos atualizados.
Subprocesso do Chrome : Módulo para abrir e fechar automaticamente uma janela do navegador Chrome, onde a previsão é visualizada.
OpenPyXL : Biblioteca para manipulação de planilhas Excel, usada para salvar e organizar os dados meteorológicos.
Descrição Funcional
Entrada de Cidade : O usuário insere o nome da cidade na interface gráfica. Caso o campo esteja vazio, uma mensagem de erro é exibida solicitando a entrada de dados.

Consulta à API do OpenWeatherMap :

A aplicação constrói uma URL de consulta com a cidade inserida e faz uma requisição HTTP.
Em caso de falha na conexão ou erro de resposta, uma mensagem de erro é exibida.
Quando a resposta é bem-sucedida, o programa extrai a temperatura (em graus Celsius) e a umidade relativa do ar da cidade.
Visualização da Previsão no Navegador :

O programa abre uma nova aba no navegador Chrome exibindo a previsão do tempo para a cidade, com uma URL de pesquisa do Google para "weather [nome da cidade]".
Após cinco segundos, o navegador é fechado automaticamente.
Salvamento dos Dados em Planilha Excel :

Os dados são registrados em uma planilha historico_clima.xlsx.
A planilha armazena os dados e a hora da consulta, cidade, temperatura, e umidade, organizados em colunas, com cabeçalho formatado e colorido para uma visualização clara.
A célula de umidade é colorida de acordo com algumas especificidades, para facilitar a identificação dos níveis de umidade.
Confirmação e Limpeza de Campos :

Após o salvamento dos dados, uma mensagem de sucesso é exibida ao usuário.
O campo de entrada é limpo para facilitar uma nova consulta, mantendo a aplicação pronta para outra pesquisa.
Funcionalidades Extras
Validação de Entrada : Garante que o usuário insira uma cidade antes de solicitar uma previsão.
Fechamento Automático do Navegador : Reduz a necessidade de interação manual para limitar a visualização da previsão.
Organização Visual na Planilha : Cada nível de umidade é colorido para facilitar a identificação de umidade muito baixa ou alta.
Benefícios do Projeto
Este projeto fornece uma maneira rápida e prática para o usuário verificar o clima em qualquer cidade, economizando tempo e facilitando o armazenamento de dados para consulta futura. A aplicação tem um design acessível, com recursos que tornam a utilização tanto pessoal quanto profissional.

Conclusão
O projeto de previsão do tempo atende seu objetivo de forma eficaz, fornecendo informações prejudiciais de forma rápida e fácil de consultar. A implementação com Tkinter, OpenWeatherMap API e OpenPyXL permite a interação com o usuário de maneira intuitiva, além de criar um histórico de dados que pode ser útil para análise.
