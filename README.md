<h1  align="center"> Envio Automatizado OnePage  </h1>

<div align="center">
<img src="https://github.com/user-attachments/assets/5ee3a97c-cf2c-48e7-84f4-b20acf3b1d5c" width="450px" />
</div>


<p align="center"> Este projeto automatiza o envio de e-mails, diariamente, com os indicadores de desempenho das lojas da Confeitaria (fictícia) Doces & Sorrisos, incluindo informações sobre faturamento (diário e anual), diversidade de produtos e ticket médio. </p>

###  :dart:Objetivo

O objetivo deste projeto é gerar e enviar automaticamente relatórios sobre o desempenho das 25 lojas da confeitaria espalhadas pelos shoppings do Brasil. 

O envio é feito por e-mail, com anexos contendo rankings de vendas, o que permite à diretoria acompanhar de forma eficiente as métricas de cada loja.

### Tecnologias Utilizadas

    Python: Linguagem principal utilizada para o desenvolvimento;

    pandas: Manipulação de dados (cálculo de indicadores de desempenho, como faturamento, ticket médio e diversidade de produtos);

    win32com: Para integração com o Microsoft Outlook e envio de e-mails automatizados;

    pathlib: Para manipulação de caminhos de arquivos.

### :wrench:Pré-requisitos

    Python 3.x (certifique-se de instalar as versões mais recentes do Python).
    
    Bibliotecas:
        pandas: pip install pandas 
        win32com (necessário para interação com o Outlook): pip install pywin32
        pathlib: Biblioteca padrão do Python.

### Arquivos Necessários

    Relatórios Excel: Os relatórios diários e anuais das lojas devem ser gerados antes de serem anexados ao e-mail. O caminho desses relatórios precisa ser configurado corretamente no código.

    Outlook Configurado: A integração com o Outlook requer que o Microsoft Outlook esteja configurado no ambiente local onde o script será executado.

### :computer:Como Usar

- Clone este repositório para sua máquina local. 

- Navegue até o diretório do projeto. 

- Abra o arquivo Automacao-de-Processo.ipynb em um ambiente Jupyter Notebook.
  
- Execute as células de código sequencialmente para carregar e analisar os dados. 

#### Configuração Inicial

> Certifique-se de ter as bibliotecas necessárias instaladas e o Outlook configurado corretamente.

> Prepare os relatórios Excel (ranking diário e anual), que serão gerados com os dados das lojas.

#### Execução do Script
> O script é executado para gerar os relatórios e enviar os e-mails de forma automatizada.

>Modifique as variáveis no código, como o caminho dos arquivos de relatórios e o destinatário do e-mail, conforme necessário.

 *Exemplo de Execução*

     import win32com.client
     import pathlib
       
1. Criar uma instância do Outlook
    1. outlook = win32com.client.Dispatch("Outlook.Application")
    2. mail = outlook.CreateItem(0)  # Criar um novo e-mail

2. Definir informações do e-mail
    1. mail.Subject = "Ranking Diário e Anual das Lojas"
    2. mail.To = "diretoria@empresa.com"
    3. mail.Body = "Prezados(as), segue em anexo o ranking diário e anual das lojas."

3. Definir o caminho dos anexos
    1. caminho_backup = pathlib.Path.cwd()
    2. dia_indicador = "2025-01-29"  # Substitua por uma variável de data real

4. Anexar arquivos
- [x] attachment1 = caminho_backup / f'{dia_indicador}_Ranking Anual.xlsx'
- [x] attachment2 = caminho_backup / f'{dia_indicador}_Ranking Dia.xlsx'
- [x] mail.Attachments.Add(str(attachment1))
- [x] mail.Attachments.Add(str(attachment2))

5.  Enviar o e-mail
      1. mail.Send()
      2. print("E-mail enviado com sucesso!")

### :exclamation:Possíveis Erros e Soluções

    "O item foi movido ou excluído."
        Esse erro ocorre quando o e-mail já foi movido ou excluído no Outlook. Certifique-se de criar um novo e-mail ao invés de tentar editar um e-mail antigo.

    "O arquivo não foi encontrado."
        Verifique se o caminho dos arquivos anexados está correto. Utilize a biblioteca pathlib para garantir que o caminho seja validado corretamente.

    "Erro de formatação no corpo do e-mail."
        Caso o corpo do e-mail não esteja sendo renderizado corretamente, revise o código HTML para garantir que as tags estão bem formatadas.

