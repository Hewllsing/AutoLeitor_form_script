# Importador de Propostas para Análise Comparativa

Este projeto automatiza o preenchimento da planilha de análise comparativa a partir de PDFs de proposta preenchidos.

O utilizador só precisa colocar os PDFs na pasta `propostas` e executar o ficheiro `Importar propostas.vbs`.

## Estrutura do Projeto


script-proposta-e-preenche-planilha
├── Importar propostas.vbs
├── guia-de-utilizacao.txt
├── modelo
│   ├── libs
│   ├── importar_pasta_propostas.py
│   ├── Analise comparativa 1.xlsx
│   └── Informações Proposta 1.pdf
├── propostas
└── resultados
    ├── Analise comparativa preenchida.xlsx
    └── nao-mexer
        ├── propostas_importadas.json
        └── resultado_importacao.txt

Como Utilizar
Preencha o PDF de proposta.

Guarde o PDF preenchido dentro da pasta:

propostas
Volte para a pasta principal do projeto.

Clique duas vezes no ficheiro:

Importar propostas.vbs
Aguarde a mensagem final de importação.

A planilha atualizada ficará em:

resultados\Analise comparativa preenchida.xlsx
O Que o Script Faz
Ao executar o importador, o sistema:

lê todos os PDFs novos dentro da pasta propostas;
extrai os campos preenchidos do formulário PDF;
preenche automaticamente a planilha de análise comparativa;
salva o resultado na pasta resultados;
evita importar o mesmo PDF mais de uma vez;
mostra uma mensagem final com o resultado da importação.
Pastas Importantes
propostas
Pasta onde devem ser colocados os PDFs preenchidos que serão importados.

modelo
Contém os ficheiros base do sistema:

planilha modelo;
PDF modelo;
script interno;
bibliotecas necessárias.
Não é recomendado alterar esta pasta sem conhecimento técnico.

resultados
Contém a planilha final preenchida.

resultados\nao-mexer
Contém ficheiros internos usados pelo sistema para controlar quais propostas já foram importadas.

Não altere nem apague esta pasta.

Requisitos
A máquina precisa ter Python 3 instalado.

As bibliotecas Python necessárias já estão incluídas dentro da pasta:

modelo\libs
Por isso, normalmente não é necessário instalar pacotes adicionais.

Como Transportar Para Outra Máquina
Para usar em outro computador, copie a pasta inteira do projeto, mantendo a mesma estrutura.

Depois, basta executar:

Importar propostas.vbs
O sistema usa caminhos relativos, então não depende do caminho específico da máquina original.

Observações
O mesmo PDF não é importado duas vezes.
A planilha modelo original não deve ser usada como resultado final.
O ficheiro final de trabalho é sempre:
resultados\Analise comparativa preenchida.xlsx
Se houver erro, verifique a mensagem exibida ao executar o importador.
Certifique-se de que os PDFs colocados em propostas estão preenchidos e salvos corretamente.
