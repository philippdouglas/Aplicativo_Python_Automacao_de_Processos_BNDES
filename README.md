# Aplicativo_Python_Automação_de_Processos

Aplicativo desenvolvido em Python, convertido para um formato executável de 64 bits compatível com Windows, com o propósito de automatizar uma série de processos na área de Comunicação e Marketing do Banco BNDES. A presente versão é uma edição parcial, abrangendo exclusivamente as funções relacionadas a informações de domínio público, provenientes de sites, blogs e redes sociais.

A função central do aplicativo reside na automação da atualização das bases de dados, formatadas como planilhas Excel. Tais atualizações feitas manualmente, consumindo um tempo considerável e com a potencialidade de posuir erros de digitação. As informações contidas nessas bases são submetidas a um processo de tratamento e organização, possibilitando a geração de relatórios e insights pertinentes às lideranças dos diversos departamentos.

O desenvolvimento do aplicativo se restringiu exclusivamente à linguagem de programação Python, valendo-se de uma ampla gama de bibliotecas, com destaque para pandas, json, BeautifulSoup, tkinter, shutil, pyautogui, requests, selenium e re, entre outras relevantes.

*Funcionalidades presentes nesta versão do aplicativo segue os menus:

Menu ABN: Efetua a atualização da base de dados contida em ABN.xlsx, mediante a obtenção de informações diretamente do endereço eletrônico: https://agenciadenoticias.bndes.gov.br.

Menu Blog: Realiza a atualização da base de dados mantida em Blog.xlsx, com dados oriundos do sítio eletrônico: https://agenciadenoticias.bndes.gov.br/blogdodesenvolvimento.

Menu Releases: Efetua a atualização da base de dados existente em Releases.xlsx, obtendo informações por meio do seguinte sítio: https://www.bndes.gov.br/wps/portal/site/home/imprensa/noticias.

Menu Seguidores: Promove a atualização da base de dados denominada Seguidores.xlsx, ao coletar o número de seguidores de distintas redes sociais, a partir da lista de links no arquivo Seguidores.xlsx.

Menu Backups: Realiza o backup completo do conteúdo da pasta ..Bases\02_Tratadas para a pasta ..Bases\05_BackUp\@Backup Tratadas, acrescentando a data e hora aos arquivos copiados.

Destaca-se que as demais funcionalidades do aplicativo foram excluídas do código, em razão da presença de informações confidenciais. Imagens e ícones foram omitidos desta versão, considerando questões de direitos autorais. Os endereços de diretórios foram ajustados para esta versão, utilizando a unidade C:. Na edição original, as pastas se encontram em uma rede, sujeitas a restrições de acesso, requerendo o emprego de credenciais de login e senha. Fora desenvolvido um sistema de autenticação que capitaliza a infraestrutura de segurança da rede, permitindo o acesso somente a usuários autorizados, com a permissão adequada para acessar as pastas contendo os arquivos de dados utilizados pela aplicação. O uso do aplicativo é viabilizado apenas após a realização do login na rede por parte desses usuários autorizados. Essa parte do código está comentada para poder rodar a aplicação localmente.

Para executar o aplicativo faça o donwload e descompacte a pasta da aplicação na unidade C: renomeie a pasta para "Automacao_DECOM&DEMKT_BNDES" , e execute o código da aplicação Automacao baseDados.py 
Caso queira testar a versão executável da aplicação execute o comando no terminal: 

```bash
pyinstaller --onefile "C:\Automacao_DECOM&DEMKT_BNDES\Automacao baseDados\Automacao baseDados.py"
```
