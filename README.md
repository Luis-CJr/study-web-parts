## As soluções focam em melhorar a gestão e acompanhamento de estudos dos usuários no SharePoint Online, utilizando as listas conteudoDeEstudos e estudosDoUsuario.

## Listas do SharePoint

**conteudoDeEstudos**
> ID: Identificador único de cada item (padrão do SharePoint).
> Title: Título do conteúdo de estudo.
> descricao: Descrição detalhada do conteúdo de estudo.

**estudosDoUsuario**
> ID: Identificador único de cada registro de estudo (padrão do SharePoint).
> Title: Nome do usuário que está realizando o estudo.
> datainicio: Data de início do estudo.
> conteudo: Referência (lookup) para conteudoDeEstudos, trazendo o título do conteúdo e o ID.
> descricao: Referência (lookup) para conteudoDeEstudos, trazendo a descrição do conteúdo e o ID.
> datafim: Data em que o estudo foi concluído.

# Funcionalidades
- Web Part 'Studyapp': Permite aos usuários registrar novos estudos, escolhendo conteúdos de uma lista predefinida e definindo datas de início.

- Web Part 'Studyapplist': Exibe uma lista dos estudos registrados, com opções para visualizar detalhes, marcar como concluído ou excluir registros.

# Tecnologias Utilizadas
> SharePoint Framework (SPFx)
> PnP JS
> jQuery
> SweetAlert2
> Bootstrap

# Pré-Requisitos
- Acesso a um ambiente SharePoint Online
- Node.js instalado localmente
- Git para clonar o repositório

# Instalação e Configuração
- Partindo da premissa que o usuário já tem as listas configuradas;
- Clone o repositório para seu ambiente local ou SharePoint Online ()
- Instale as dependencias do projeto ("npm install")
- Gere o pacote com a solução (
    > gulp bundle --ship  
    > gulp package-solution --ship
  ), isso irá gerar um pacote '.sppkg' na pasta sharepoint/solution
- Vá para o App Catalog do Sharepoint e carregue o arquivo '.sppkg'
- Disponibilize a solução implantada no seu ambiente indo até a página Conteúdos do Site: **seu-sharepoint.com/_layouts/15/viewlsts.aspx**
- Selecione "Novo"
- Selecione "Aplicativo"
- Estando na página **seu-sharepoint.com/_layouts/15/appStore.aspx**, adicione a sua solução
