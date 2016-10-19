# <a name="image-callouts-word-add-in-sample:-load,-edit,-and-insert-images"></a>Exemplo de suplemento do Word da imagem: carregar, editar e inserir imagens

**Sumário**

* [Resumo](#summary)
* [Ferramentas necessárias](#required-tools)
* [Como instalar certificados](#how-to-install-certificates)
* [Como configurar e executar o aplicativo](#how-to-set-up-and-run-the-app)
* [Como executar o suplemento no Word 2016 para Windows](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [Perguntas frequentes](#faq)
* [Perguntas e comentários](#questions-and-comments)
* [Saiba mais](#learn-more)


## <a name="summary"></a>Resumo

Este exemplo de suplemento do Word mostra como:

1. Criar um suplemento do Word com Typescript.
2. Carregar imagens do documento para o suplemento.
3. Editar imagens no suplemento, usando a API de tela em HTML e inserir as imagens em um documento do Word.
4. Implementar comandos de suplemento que iniciam um suplemento na faixa de opções e executam um script a partir da faixa de opções e de um menu de contexto.
5. Use o Office UI Fabric para criar uma experiência nativa parecida com o Word para o suplemento.

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

Definição - **comando de suplemento**: uma extensão para a interface do usuário do Word que permite iniciar o suplemento no painel de tarefas ou executar um script a partir da faixa de opções ou de um menu de contexto.

Se você quiser ver isso em ação, passe diretamente para [configuração do Word 2016 para Windows](#word-2016-for-windows-set-up) e use esse [manifesto](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml).

## <a name="required-tools"></a>Ferramentas necessárias

Para usar o exemplo de suplemento do Word Textos explicativos da imagem, são necessários.

* Word 2016 16.0.6326.0000, ou superior, ou qualquer cliente que seja compatível com a API Javascript do Word. Este exemplo faz uma verificação de requisitos para ver se ele está sendo executado em um host compatível com as APIs JavaScript.
* npm (https://www.npmjs.com/) para instalar as dependências. O suplemento vem com [NodeJS](https://nodejs.org/en/).
* (Windows) [Git Bash](http://www.git-scm.com/downloads).
* Clone este repositório em seu computador local.

> Observação: O Word para Mac 2016 não é compatível com comandos de suplemento neste momento. Esse exemplo pode ser executado no Mac sem os comandos de suplemento.

## <a name="how-to-install-certificates"></a>Como instalar certificados

Será necessário um certificado para executar esse exemplo já que os comandos de suplemento exigem HTTPS e, como eles não têm uma interface do usuário, você não pode aceitar certificados inválidos. Execute [./gen-cert.sh](#gen-cert.sh) para criar o certificado e, em seguida, será necessário instalar ca.crt em sua loja de Autoridades de Certificação Raiz Confiáveis (Windows).

## <a name="how-to-set-up-and-run-the-app"></a>Como configurar e executar o aplicativo

1. Instale o Gerenciador de definição do TypeScript digitando ```npm install typings -g``` na linha de comando.
2. Instale as definições do Typescript identificadas em typings.json executando ```typings install``` no diretório raiz do projeto na linha de comando.
3. Instale as dependências do projeto identificadas em package.json executando ```npm install``` no diretório raiz do projeto.
4. Instale o ```npm install -g gulp``` do Gulp.
5. Copie os arquivos do Fabric e do JQuery executando o ```gulp copy:libs```. (Windows) Se você tiver um problema aqui, verifique se *%APPDATA%\npm* está em sua variável path.
6. Execute a tarefa padrão do Gulp executando ```gulp``` a partir do diretório raiz do projeto. Se as definições do TypeScript não estiverem atualizadas, você receberá um erro aqui.

Nesse momento, esse suplemento do exemplo foi implementado. Agora, você precisa informar ao Word onde encontrar o suplemento.

### <a name="word-2016-for-windows-set-up"></a>Configuração do Word 2016 para Windows

1. (Windows) Descompacte e execute essa [chave de registro](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark) para ativar o recurso de comandos de suplemento. Isso é necessário porque os comandos de suplemento são um **recurso de visualização**.
2. Crie um compartilhamento de rede ou [compartilhe uma pasta para a rede](https://technet.microsoft.com/en-us/library/cc770880.aspx) e coloque o arquivo de manifesto [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) nele.
3. Inicie o Word e abra um documento.
4. Escolha a guia **Arquivo** e escolha **Opções**.
5. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
6. Escolha **Catálogos de Suplementos Confiáveis**.
7. Na caixa **URL do Catálogo**, digite o caminho de rede para o compartilhamento de pasta que contém manifest-word-add-in-canvas.xml e escolha **Adicionar Catálogo**.
8. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.
9. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office. Feche e reinicie o Word.

## <a name="how-to-run-the-add-in-in-word-2016-for-windows"></a>Como executar o suplemento no Word 2016 para Windows

1. Abra um documento do Word.
2. Na guia **Inserir** no Word 2016, escolha **Meus Suplementos**.
3. Selecione a guia **Pasta compartilhada**.
4. Escolha o **suplemento Texto explicativo da imagem** e **Inserir**.
5. Se os comandos de suplemento forem compatíveis com sua versão do Word, a interface do usuário informará que o suplemento foi carregado. Você pode usar a guia do **suplemento Texto explicativo** para carregar o suplemento na interface do usuário e para inserir uma imagem no documento. Você também pode usar o menu de contexto do botão direito do mouse para inserir uma imagem no documento.
6. O suplemento será carregado no painel de tarefas se os comandos de suplemento não forem compatíveis com sua versão do Word. Você precisará inserir uma imagem no documento do Word para usar a funcionalidade do suplemento.
7. Selecione uma imagem no documento do Word e carregue-a no painel de tarefas escolhendo *Carregar imagem a partir do doc*. Agora você pode inserir os textos explicativos na imagem. Escolha *Inserir imagem no documento* para posicionar a imagem atualizada no documento do Word. O suplemento gerará descrições de espaço reservado para cada um dos textos explicativos.

## <a name="faq"></a>Perguntas frequentes

* Os comandos de suplemento funcionarão no Mac e no iPad? Não, eles não funcionarão no Mac ou no iPad desde a publicação deste arquivo Leiame.
* Por que meu suplemento não é exibido na janela **Meus Suplementos**? Pode haver um erro no manifesto de seu suplemento. Sugerimos que você valide o manifesto no [esquema do manifesto](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD).
* Por que o arquivo de função não é chamado para meus comandos de suplemento? Os comandos de suplemento exigem HTTPS. Como os comandos de suplemento exigem TLS e não existe uma interface do usuário, não é possível ver se há um problema de certificado. Se você tiver que aceitar um certificado inválido no painel de tarefas, o comando de suplemento não funcionará.
* Por que os comandos de instalação do npm travam? Eles provavelmente não estão travados. Eles demoram um pouco para iniciar no Windows.

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre o exemplo de suplemento do Word Texto explicativo da imagem. Você pode enviar perguntas e sugestões na seção [Problemas](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues) deste repositório.

As perguntas sobre o desenvolvimento de suplementos em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Não deixe de marcar as perguntas ou comentários com [office-js], [word-addins] e [API]. Estamos observando essas marcas.

## <a name="learn-more"></a>Learn more

Veja mais recursos para ajudá-lo a criar suplementos baseados na API Javascript do Word:

* [Exemplos e documentação de suplementos do Word](https://dev.office.com/word)
* [Exemplo de histórias engraçadas](https://github.com/OfficeDev/Word-Add-in-SillyStories) – Saiba como carregar arquivos docx a partir de um serviço e inserir os arquivos em um documento aberto do Word.
* [Exemplo de Autenticação de Servidor de Suplemento do Office para Node.js](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) - saiba como usar os provedores OAuth do Azure e do Google para autenticar usuários de suplementos.

## <a name="copyright"></a>Direitos autorais
Copyright © 2016 Microsoft. Todos os direitos reservados.
