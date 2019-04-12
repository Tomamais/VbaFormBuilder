# VbaFormBuilder

Aplicativo inspirado no Modelo de Cadastro

*Requisitos*

+ Macros ativadas (óbvio)
+ Confiança no projeto VBA ativado
+ Referências às bibliotecas:
	- Microsoft Visual Basic for Applications Extensibility 5.3
	- Microsoft Office 16.0 Access database engine Object
	- Microsoft Scripting Runtime

*Compatibilidade*

Este aplicativo foi criado no Excel 2016 e testado no Excel 2013 (32 e 64 bits). O suporte oficial é focado na versão 2016 e superior. O funcionamento do aplicativo em versões anteriores é por conta e risco.

*Uso*

* Abra o arquivo VbaFormBuilder.xlsm na raiz do repositório ativando as macros
* Ative a confiança no Projeto VBA
* Clique no botão Importar Código e aguarde a finalização do processo
* Clique em Iniciar
* Siga as instruções deste vídeo:
 
[![Alt text](https://img.youtube.com/vi/Wry1AWqUX0E/0.jpg)](https://www.youtube.com/watch?v=Wry1AWqUX0E)

*Melhorias para a branch melhorias_Q2_2019*

+ No form de pesquisa (checar o de cadastro), usar a propriedade Value ao invés do text - OK
+ Permitir selecionar mais de uma tabela a gerar formulários
	- Transformar o comboBox o comboBox de tabelas em um listbox multiselecao (ideal auto-selecionar as tabelas com relacao direta) - OK
	- Adicionar função LoadDependentCombos - OK
	- Adicionar classes dependentes na função acima (limpeza no queryclose idem) - OK
	- Mesclar modTypes para todos os formulários gerados - OK
+ Permitir configurar ID e VALUE para campos combobox tanto no cadastro como na pesquisa - OK
+ Botões de navegação opcionais - OK
+ Tratar campos de chave primária sem auto-numeração
+ Remover _ dos labels em formulários (baixa prioridade)