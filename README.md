📋 Descrição do Projeto

Sistema completo de gestão de recibos para oficinas mecânicas, desenvolvido em Python com interface gráfica PyQt5. Permite criar, editar, visualizar e imprimir recibos com integração direta para geração de PDFs profissionais.
✨ Funcionalidades Principais

    Gestão completa de recibos: Criação, edição, busca e exclusão de recibos

    Interface moderna: Design atualizado com validação de dados em tempo real

    Geração de PDF: Criação automática de recibos em formato PDF com layout profissional

    Armazenamento em Excel: Todos os dados são salvos em planilha Excel para fácil backup

    Consulta de CEP: Integração com API ViaCEP para preenchimento automático de endereços

    Cálculos automáticos: Sistema de cálculos de valores com descontos por item

    Gestão de peças e serviços: Controle detalhado de itens com tipos, quantidades e valores

🛠️ Tecnologias Utilizadas

    Python 3.7+: Linguagem de programação principal

    PyQt5: Framework para interface gráfica moderna

    Pandas: Manipulação e armazenamento de dados em Excel

    WeasyPrint: Geração de PDFs a partir de templates HTML

    Jinja2: Template engine para formatação de documentos

    Requests: Comunicação com APIs externas (ViaCEP)

    Pillow: Processamento de imagens para logos

📦 Instalação e Configuração
Pré-requisitos

    Python 3.7 ou superior instalado

    Pip (gerenciador de pacotes do Python)

Passos de Instalação

    Instale as dependências:

bash

pip install -r requirements.txt

    Estrutura de pastas:
    
   <img width="753" height="163" alt="image" src="https://github.com/user-attachments/assets/de0f6097-553c-4644-8fa9-c8779433c0a9" />

    Configure o logo da oficina (opcional):

        Coloque uma imagem chamada logo.png na pasta principal

        Dimensões recomendadas: 100x100 pixels

🖥️ Como Utilizar o Sistema
1. Iniciando a Aplicação

Execute o arquivo principal:
bash

python main.py

2. Criando um Novo Recibo

    Dados do Recibo: O sistema gera automaticamente um número sequencial

    Dados do Cliente: Preencha todas as informações do cliente

        O CEP será automaticamente completado com dados da ViaCEP

        Telefone e CPF/CNPJ são formatados automaticamente

    Dados do Veículo: Informações completas do veículo

        Quilometragem de entrada e saída

        Combustível e box são selecionados via combobox

    Itens e Serviços: Adicione peças e serviços com:

        Tipo (Peça, Serviço)

        Código/referência

        Descrição detalhada

        Valor unitário e quantidade

        Percentual de desconto por item

    Informações Finais:

        Responsável pelo serviço

        Situação atual

        Condições de pagamento

        Observações gerais

        Próxima revisão recomendada

3. Gerenciando Recibos Existentes

    Buscar Recibo: Digite o número do recibo no campo de busca

    Editar Recibo: Após buscar, faça as alterações necessárias

    Excluir Recibo: Use o botão "Deletar Recibo Atual" (com confirmação)

    Salvar Alterações: Sempre clique em "Salvar Recibo"

4. Gerando PDF

    Clique em "Gerar e Visualizar Recibo (PDF)" para criar um documento impresso

    O PDF será aberto automaticamente no visualizador padrão

    Os arquivos são salvos na pasta Recibos_Gerados/ com numeração automática

📊 Estrutura do Arquivo Excel

O sistema utiliza um arquivo Excel (Recibos_Historico.xlsx) com a seguinte estrutura:

<img width="502" height="357" alt="image" src="https://github.com/user-attachments/assets/1e4b4e72-463e-49ef-8f73-1fe4da3dddfc" />
<img width="509" height="335" alt="image" src="https://github.com/user-attachments/assets/4759224e-7da6-46c5-9091-735f4b8674d9" />
<img width="507" height="329" alt="image" src="https://github.com/user-attachments/assets/0e6aab41-623a-4693-b5c9-3a8dbc02b6e4" />
<img width="509" height="415" alt="image" src="https://github.com/user-attachments/assets/bbf28636-87b4-4809-a8ca-0e923923ecbe" />
🎨 Personalização
Modificando o Template do PDF

Edite o arquivo recibo_template.html para alterar o layout do PDF gerado.
Alterando Informações da Oficina

Modifique a constante INFO_OFICINA no código fonte para atualizar:

    Nome da oficina

    Endereço completo

    CNPJ

    Telefones de contato

    E-mail

🔧 Solução de Problemas
Problemas Comuns

    Erro ao gerar PDF:

        Verifique se o WeasyPrint está instalado corretamente

        Confirme que o template HTML existe e está acessível

    Falha na consulta de CEP:

        Verifique a conexão com internet

        Confirme se o serviço ViaCEP está disponível

    Erro ao salvar no Excel:

        Feche o arquivo Excel se estiver aberto em outro programa

        Verifique as permissões de escrita na pasta

Logs e Debug

    Os logs detalhados são exibidos no console durante a execução

    Erros são registrados com timestamp para facilitar troubleshooting

📝 Licença

Este projeto é destinado para uso interno de oficinas mecânicas.
🤝 Suporte

Para reportar bugs ou sugerir melhorias:

    Verifique a documentação existente

    Consulte os logs de erro no console

    Entre em contato com a equipe de desenvolvimento

🔄 Atualizações Futuras

    Integração com sistema de estoque

    Controle de usuários e permissões

    Relatórios gerenciais e analytics

    Backup em nuvem automático

    Versão mobile para consulta rápida

Nota: Este sistema foi desenvolvido para otimizar o fluxo de trabalho em oficinas mecânicas, substituindo processos manuais por uma solução digital integrada para gestão de recibos.
