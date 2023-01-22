# RPA_Envio_Whats V1.0

Aplicação RPA desenvolvida em Python com a finalidade de leitura de um arquivo Excel e envio de mensagens pré-progamadas alterando seus principais dados via WhatsApp.

Neste projeto foi utilizado as bibliotecas: 
pywhatkil
keyboard
time 
datatime
win32.com (atributos client e Dispath)

Funcionamento: 

ATENÇÂO: Após a execução do RPA não é permitido a utilização da máquina até o fim do processo.
A aplicação abre o arquivo Excel no diretório definido no código e após isso realiza uma leitura dos dados definidos nas colunas, por padrão é necessário o preenchimento das colunas Aluno, Valor e Número de telefone (Necessário ser cadastrado no WhatsApp).
O RPA coletando esses dados linha por linha abre o navegador padrão do seu computar na página do WhatsApp Web (Necessário estar conectado o seu dispositivo e ambos com conexão na internet), feito esses procedimentos o RPA inicia uma conversa com o número presente na planilha e digita no campo de mensagem os dados marcados no código. Após todos os dados preenchidos é realizado o envio da mensagem para o usuário, o navegador é finalizado e ele prossegue para a próxima linha realizando em Loop até encontrar uma linha em branco. 

Adicionais:
Será implementado novas funcionalidades e uma interface ao usuário para melhor utilização de seus usuários. 
