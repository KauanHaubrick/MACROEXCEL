# MACROEXCEL
Script em VBA criado para utilizar um MACRO no Excel com o intuito de automatizar a busca dos pesos das ordens de venda no SAP.<br>
Exemplo do funcionamento do SCRIPT.<br>
Ele busca o valor da celular A4 (Numero da OV),<br>
Com o SAP aberto na tela inicial, ele começa a operar abrindo a transação VA03,<br>
Na VA03 ele cola o valor da celular A4 e acessa a OV,<br>
No campo VALOR TOTAL ele copia o texto que contem o peso da OV e retorna com o valor para o EXCEL,<br>
Ele cola o valor na celula D4 e continua o processo até encontrar a celula vazia ou chegar na celula limite que está definda.<br>
