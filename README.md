# automacao-lembrete-vacinas
 Automação em Python para o lembrete de vacinas vencendo em uma  clínica veterinária

- A tabela contém: um cabeçalho de 3 linhas, colunas Clientes, Telefones, Animal, Vacina. Todas as vacinas da tabela são supostas a vencer na semana;

- O código obtém o diretório do script e supõe que a tabela estará junto;
- Conversão para CSV;
- Mapeamento com dicionário para refinar o envio final (por exemplo: vacina antirrábica, na mensagem final envia 'vacina de raiva do xxx' ao cliente);
- Função de envio de mensagens:
    - Simular: para testes; printa as mensagens e o destinatário no console ao invés de enviá-las;
    - Todos os nomes contidos na mensagem são ajustados, capitalizados etc para melhor resultado final;
    - Obtém primeiro nome do cliente, capitalizado;
    - Algumas linhas de Telefones contém mais de um número, separado por vírgulas. O código prioriza o primeiro número de celular disponível(não fixo);
    - Caso não haja nenhum, ele printa um erro e segue ao próximo cliente;
    - Agrupamento por clientes e animais. Evita envio de múltipla mensagens caso o cliente tenha mais de um animal com vacinas vencendo, ou múltiplas vacinas de um animal;
    - A mensagem final discerne produtos(ex: vermífugo, antiparasitário) que podem ser aplicados em casa de vacinas, no primeiro caso perguntando da aplicação para que possamos atualizar o sistema, no segundo perguntando se deseja agendar um horário para renovar a vacina;
    - Envio direto no whatsapp Web usando pywhatkit;
    - Envio automático de mensagens e fechamento da aba antes de abrir a próxima, até o fim da execução.
 
 
