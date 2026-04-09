const STORAGE_KEY = "planner-treinos:v3";
const LEGACY_STORAGE_KEYS = [STORAGE_KEY, "planner-treinos:v2", "planner-treinos:v1"];
const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const createExerciseCatalog = (entries) => Object.freeze(entries.map(([
  name,
  level,
  type,
  equipment,
  secondary,
  description,
  postureTips,
  variations,
  goals,
  aliases = [],
]) => ({
  name,
  level,
  type,
  equipment,
  secondary,
  description,
  postureTips,
  variations,
  goals,
  aliases,
})));

const EXERCISE_LIBRARY = Object.freeze({
  Peito: createExerciseCatalog([
    ["Supino reto com barra", "Intermediario", "Composto", "Barra e banco", ["Triceps", "Deltoide anterior"], "Pressao horizontal classica para ganho global de peitoral.", "Mantenha escapas retraidas e os pes firmes no solo.", ["Supino reto com halteres", "Chest press na maquina"], ["Hipertrofia", "Forca"], ["Supino reto"]],
    ["Supino inclinado com barra", "Intermediario", "Composto", "Barra e banco inclinado", ["Deltoide anterior", "Triceps"], "Enfatiza a porcao clavicular do peitoral.", "Evite perder a curva neutra da lombar no banco inclinado.", ["Supino inclinado com halteres", "Smith inclinado"], ["Hipertrofia", "Forca"], ["Supino inclinado"]],
    ["Supino declinado com barra", "Intermediario", "Composto", "Barra e banco declinado", ["Triceps", "Deltoide anterior"], "Trabalha peito com maior enfase em fibras inferiores.", "Desca controlado e barra alinhada ao esterno inferior.", ["Supino declinado com halteres"], ["Hipertrofia"], ["Supino declinado"]],
    ["Supino reto com halteres", "Iniciante", "Composto", "Halteres e banco", ["Triceps", "Deltoide anterior"], "Permite maior amplitude e simetria entre os lados.", "Controle a descida e nao bata os halteres no topo.", ["Supino neutro com halteres", "Squeeze press"], ["Hipertrofia", "Resistencia"]],
    ["Supino inclinado com halteres", "Iniciante", "Composto", "Halteres e banco inclinado", ["Deltoide anterior", "Triceps"], "Variante livre para peitoral superior.", "Mantenha os punhos neutros e cotovelos ligeiramente abaixo dos ombros.", ["Supino inclinado alternado"], ["Hipertrofia", "Resistencia"]],
    ["Crucifixo reto com halteres", "Iniciante", "Isolado", "Halteres e banco", ["Deltoide anterior"], "Alongamento controlado para peitoral em plano reto.", "Pense em abracar um barril e evite estender demais os cotovelos.", ["Crucifixo no cabo", "Peck deck"], ["Hipertrofia"], ["Crucifixo reto"]],
    ["Crucifixo inclinado com halteres", "Intermediario", "Isolado", "Halteres e banco inclinado", ["Deltoide anterior"], "Trabalha aducao horizontal com foco em peitoral superior.", "Desca lenta e cotovelos sem travar na subida.", ["Crossover na polia baixa"], ["Hipertrofia"], ["Crucifixo inclinado"]],
    ["Crossover na polia alta", "Iniciante", "Isolado", "Polia", ["Deltoide anterior"], "Mantem tensao continua no peitoral durante todo o arco.", "Fixe o tronco e conduza as maos ate a linha media do corpo.", ["Crossover na polia media", "Crossover unilateral"], ["Hipertrofia", "Emagrecimento"], ["Cross over"]],
    ["Crossover na polia baixa", "Intermediario", "Isolado", "Polia", ["Deltoide anterior"], "Enfase maior no feixe superior do peitoral.", "Movimento ascendente controlado sem compensar com lombar.", ["Crucifixo inclinado com halteres"], ["Hipertrofia"]],
    ["Peck deck", "Iniciante", "Isolado", "Maquina", ["Deltoide anterior"], "Opcao estavel para isolamento do peitoral.", "Ajuste o banco para alinhar maos ao centro do peito.", ["Crucifixo na maquina convergente"], ["Hipertrofia", "Iniciante"], ["Peck deck", "Voador"]],
    ["Chest press na maquina", "Iniciante", "Composto", "Maquina", ["Triceps", "Deltoide anterior"], "Empurrada guiada para construir tecnica e volume.", "Mantenha escapulas apoiadas e nao trave os cotovelos.", ["Chest press convergente", "Supino no smith"], ["Hipertrofia", "Iniciante"]],
    ["Flexao de bracos tradicional", "Iniciante", "Composto", "Peso corporal", ["Triceps", "Core"], "Movimento basico de empurrar com controle corporal.", "Corpo alinhado da cabeca ao calcanhar durante toda a repeticao.", ["Flexao inclinada", "Flexao com apoio nos joelhos"], ["Resistencia", "Emagrecimento"], ["Flexao de bracos"]],
    ["Flexao inclinada", "Iniciante", "Composto", "Peso corporal e apoio", ["Triceps", "Core"], "Versao regressiva para dominar o padrao de empurrar.", "Use um apoio firme e mantenha o tronco rigido.", ["Flexao na parede", "Flexao no banco"], ["Iniciante", "Resistencia"]],
    ["Paralelas com foco em peito", "Avancado", "Composto", "Barras paralelas", ["Triceps", "Deltoide anterior"], "Exercicio intenso com grande alongamento do peitoral.", "Incline levemente o tronco e desca com controle.", ["Paralelas assistidas"], ["Hipertrofia", "Forca"]],
    ["Squeeze press com halteres", "Intermediario", "Isolado", "Halteres e banco", ["Triceps"], "Mantem aducao constante apertando os halteres juntos.", "Pressione um halter contra o outro durante toda a subida.", ["Hex press"], ["Hipertrofia"]],
  ]),
  Costas: createExerciseCatalog([
    ["Puxada frontal pronada", "Iniciante", "Composto", "Polia e barra longa", ["Biceps", "Redondo maior"], "Puxada vertical para largura de dorsais.", "Leve a barra ao peitoral sem perder a postura do tronco.", ["Puxada frontal aberta", "Puxada neutra"], ["Hipertrofia", "Resistencia"], ["Puxada frontal"]],
    ["Puxada frontal supinada", "Iniciante", "Composto", "Polia e barra reta", ["Biceps", "Braquial"], "Variante com maior participacao dos flexores do cotovelo.", "Puxe com cotovelos para baixo e peito elevado.", ["Pulldown supinado"], ["Hipertrofia"], ["Puxada supinada"]],
    ["Puxada neutra na maquina", "Iniciante", "Composto", "Maquina", ["Biceps", "Trapezio medio"], "Trajetoria guiada para dorsais e controle de escapulas.", "Nao eleve os ombros no inicio da puxada.", ["Puxada neutra na polia"], ["Hipertrofia", "Iniciante"]],
    ["Remada curvada com barra", "Intermediario", "Composto", "Barra", ["Biceps", "Lombar"], "Movimento base para espessura dorsal.", "Incline o tronco mantendo coluna neutra e barra proxima ao corpo.", ["Pendlay row", "Remada curvada pronada"], ["Hipertrofia", "Forca"], ["Remada curvada"]],
    ["Remada unilateral com halter", "Iniciante", "Composto", "Halter e banco", ["Biceps", "Trapezio medio"], "Excelente para corrigir assimetrias entre os lados.", "Apoie bem a mao livre e puxe o cotovelo para tras.", ["Remada serrote"], ["Hipertrofia"], ["Remada serrote"]],
    ["Remada baixa na polia", "Iniciante", "Composto", "Polia", ["Biceps", "Romboides"], "Mantem tensao constante na fase concentrica e eccentrica.", "Evite jogar o tronco para tras e inicie com depressao escapular.", ["Remada baixa com triangulo"], ["Hipertrofia", "Resistencia"], ["Remada baixa"]],
    ["Remada cavalinho", "Intermediario", "Composto", "Barra T", ["Biceps", "Trapezio medio"], "Variante potente para densidade de costas.", "Tronco firme e movimento liderado pelos cotovelos.", ["T-bar row apoiada"], ["Hipertrofia", "Forca"]],
    ["Pulldown na polia", "Iniciante", "Isolado", "Polia", ["Triceps longo", "Peitoral"], "Extensao de ombro focada em dorsais.", "Mantenha bracos semi estendidos e ombros longe das orelhas.", ["Straight-arm pulldown"], ["Hipertrofia", "Emagrecimento"], ["Pulldown"]],
    ["Barra fixa pronada", "Avancado", "Composto", "Peso corporal", ["Biceps", "Core"], "Exercicio classico para largura e controle corporal.", "Inicie pela escapula e evite impulso das pernas.", ["Barra fixa assistida"], ["Forca", "Hipertrofia"], ["Barra fixa"]],
    ["Barra fixa supinada", "Avancado", "Composto", "Peso corporal", ["Biceps", "Braquial"], "Versao com maior recrutamento de biceps.", "Suba com peito aberto e pes cruzados para estabilidade.", ["Chin-up assistido"], ["Forca", "Hipertrofia"]],
    ["Levantamento terra convencional", "Avancado", "Composto", "Barra", ["Gluteos", "Posterior de coxa"], "Padrao de dobradica completo com forte demanda de cadeia posterior.", "Barra rente ao corpo e lombar neutra desde a partida.", ["Terra com hex bar", "Terra romeno"], ["Forca", "Hipertrofia"], ["Levantamento terra"]],
    ["Remada apoiada no banco inclinado", "Iniciante", "Composto", "Halteres e banco inclinado", ["Biceps", "Romboides"], "Reduz compensacoes e facilita foco na dorsal.", "Mantenha o peito apoiado no banco para estabilizar o tronco.", ["Seal row"], ["Hipertrofia"]],
    ["Remada maquina articulada", "Iniciante", "Composto", "Maquina", ["Biceps", "Trapezio medio"], "Trajetoria estavel para acumular volume com seguranca.", "Ajuste o assento para puxar na direcao do abdome superior.", ["Remada convergente"], ["Hipertrofia", "Iniciante"]],
    ["Pull-over na polia", "Intermediario", "Isolado", "Polia", ["Peitoral", "Triceps longo"], "Mantem tensao continua na extensao do ombro.", "Use arco controlado e nao dobre demais os cotovelos.", ["Pull-over com halter"], ["Hipertrofia"], ["Pull over"]],
    ["Face pull para costas altas", "Iniciante", "Isolado", "Polia e corda", ["Deltoide posterior", "Rotadores externos"], "Excelente para postura e parte alta das costas.", "Puxe a corda na direcao do rosto com cotovelos altos.", ["Face pull sentado"], ["Resistencia", "Reabilitacao"], ["Face pull"]],
  ]),
  Pernas: createExerciseCatalog([
    ["Agachamento livre com barra", "Intermediario", "Composto", "Barra", ["Gluteos", "Core"], "Principal exercicio global para membros inferiores.", "Desca com joelhos alinhados aos pes e tronco firme.", ["Agachamento high bar", "Agachamento low bar"], ["Hipertrofia", "Forca"], ["Agachamento livre"]],
    ["Agachamento frontal com barra", "Avancado", "Composto", "Barra", ["Core", "Gluteos"], "Aumenta a demanda de quadriceps e controle postural.", "Cotovelos altos e tronco mais vertical durante todo o movimento.", ["Front squat no smith"], ["Forca", "Hipertrofia"]],
    ["Hack squat", "Iniciante", "Composto", "Maquina", ["Gluteos", "Posterior de coxa"], "Opcao guiada para gerar alto volume de pernas.", "Mantenha quadril e costas apoiados na plataforma.", ["Hack reverse"], ["Hipertrofia", "Iniciante"], ["Agachamento hack"]],
    ["Leg press 45", "Iniciante", "Composto", "Maquina", ["Gluteos", "Posterior de coxa"], "Permite alta sobrecarga com estabilidade.", "Nao solte o quadril no final da descida.", ["Leg press horizontal", "Leg press unilateral"], ["Hipertrofia", "Forca"], ["Leg press"]],
    ["Cadeira extensora", "Iniciante", "Isolado", "Maquina", ["Reto femoral"], "Isolamento de quadriceps com facil controle de carga.", "Alinhe o eixo da maquina ao joelho e suba sem impulsos.", ["Extensora unilateral"], ["Hipertrofia", "Reabilitacao"]],
    ["Mesa flexora", "Iniciante", "Isolado", "Maquina", ["Gastrocnemio"], "Foco em flexores de joelho e posterior de coxa.", "Mantenha quadril colado ao banco e retorno controlado.", ["Flexora sentada", "Flexora unilateral"], ["Hipertrofia", "Reabilitacao"]],
    ["Agachamento bulgaro com halteres", "Intermediario", "Composto", "Halteres e banco", ["Gluteos", "Core"], "Unilateral intenso para quadriceps e gluteos.", "Desca em linha reta sem deixar o joelho colapsar.", ["Bulgaro no smith"], ["Hipertrofia", "Resistencia"], ["Agachamento bulgaro"]],
    ["Afundo caminhando com halteres", "Intermediario", "Composto", "Halteres", ["Gluteos", "Posterior de coxa"], "Desenvolve pernas com demanda alta de equilibrio.", "Passos firmes e joelho de tras proximo ao solo.", ["Afundo estacionario"], ["Hipertrofia", "Emagrecimento"], ["Afundo"]],
    ["Passada no smith", "Iniciante", "Composto", "Smith machine", ["Gluteos", "Core"], "Versao guiada para controlar melhor a trajetoria.", "Mantenha a passada longa e o tronco estavel.", ["Passada reversa no smith"], ["Hipertrofia"], ["Passada"]],
    ["Stiff com barra", "Intermediario", "Composto", "Barra", ["Gluteos", "Lombar"], "Dobradiça de quadril com foco em posterior de coxa.", "Barra rente as pernas e joelhos levemente flexionados.", ["Stiff com halteres"], ["Hipertrofia", "Forca"], ["Stiff"]],
    ["Terra romeno com halteres", "Intermediario", "Composto", "Halteres", ["Gluteos", "Lombar"], "Variante livre para cadeia posterior com amplitude controlada.", "Quadril vai para tras sem arredondar a coluna.", ["Terra romeno unilateral"], ["Hipertrofia", "Resistencia"], ["Terra romeno"]],
    ["Agachamento sumo com halter", "Iniciante", "Composto", "Halter", ["Gluteos", "Adutores"], "Base ampla para trabalhar pernas e adutores.", "Joelhos acompanham a abertura dos pes durante a descida.", ["Agachamento sumo com barra"], ["Hipertrofia", "Emagrecimento"]],
    ["Step-up com halteres", "Iniciante", "Composto", "Halteres e caixa", ["Gluteos", "Panturrilhas"], "Movimento unilateral funcional para pernas.", "Apoie o pe inteiro na caixa e evite impulso exagerado da perna de baixo.", ["Step-up alto", "Step-up lateral"], ["Resistencia", "Emagrecimento"], ["Step-up"]],
    ["Avanco reverso com halteres", "Iniciante", "Composto", "Halteres", ["Gluteos", "Core"], "Opcao amigavel para joelhos com grande controle.", "Leve a perna para tras mantendo o tronco estavel.", ["Avanco alternado"], ["Hipertrofia", "Resistencia"]],
    ["Sissy squat assistido", "Avancado", "Isolado", "Peso corporal e apoio", ["Reto femoral"], "Variante exigente para quadriceps com grande alavanca.", "Use apoio seguro e mantenha quadril estendido durante a descida.", ["Spanish squat"], ["Hipertrofia"]],
  ]),
  Gluteos: createExerciseCatalog([
    ["Elevacao pelvica com barra", "Intermediario", "Composto", "Barra e banco", ["Posterior de coxa", "Core"], "Principal exercicio para gluteo maximo com alta sobrecarga.", "Queixo recolhido e costelas baixas no topo do movimento.", ["Hip thrust no smith", "Hip thrust unilateral"], ["Hipertrofia", "Forca"], ["Elevacao pelvica"]],
    ["Ponte glutea unilateral", "Iniciante", "Composto", "Peso corporal", ["Core", "Posterior de coxa"], "Boa opcao para estabilidade de quadril e ativacao unilateral.", "Empurre o solo pelo calcaneo e mantenha quadris alinhados.", ["Ponte glutea bilateral"], ["Resistencia", "Reabilitacao"], ["Ponte glutea"]],
    ["Coice no cabo", "Iniciante", "Isolado", "Polia", ["Posterior de coxa"], "Isola extensao de quadril com tensao continua.", "Tronco firme e sem girar a pelve durante a extensao.", ["Coice na maquina"], ["Hipertrofia"], ["Coice no cabo"]],
    ["Abducao de quadril na maquina", "Iniciante", "Isolado", "Maquina", ["Gluteo medio"], "Foco em gluteo medio e estabilidade do quadril.", "Evite jogar o tronco para frente para roubar a carga.", ["Abducao com miniband"], ["Hipertrofia", "Reabilitacao"], ["Abducao de quadril"]],
    ["Kickback na maquina", "Iniciante", "Isolado", "Maquina", ["Posterior de coxa"], "Extensao guiada de quadril com facil controle.", "Mantenha a pelvis neutra e finalize com contracao voluntaria.", ["Kickback com caneleira"], ["Hipertrofia"]],
    ["Agachamento sumo com barra", "Intermediario", "Composto", "Barra", ["Adutores", "Quadriceps"], "Base ampla para fortalecer gluteos e adutores.", "Pense em abrir o chao com os pes ao subir.", ["Agachamento sumo com halter"], ["Hipertrofia", "Forca"]],
    ["Step-up alto", "Intermediario", "Composto", "Halteres e caixa alta", ["Quadriceps", "Panturrilhas"], "Variante com maior exigencia de extensao de quadril.", "Suba controlando a descida e sem impulso da perna de tras.", ["Step-up com joelho alto"], ["Hipertrofia", "Resistencia"]],
    ["Glute bridge com miniband", "Iniciante", "Composto", "Peso corporal e miniband", ["Gluteo medio", "Core"], "Ativacao acessivel para gluteos e controle de joelhos.", "Empurre a banda para fora sem perder a ponte.", ["Ponte com pausa isometrica"], ["Iniciante", "Reabilitacao"]],
    ["Stiff com foco em gluteos", "Intermediario", "Composto", "Barra ou halteres", ["Posterior de coxa", "Lombar"], "Dobradiça de quadril enfatizando extensao glutea.", "Segure a carga proxima do corpo e empurre o quadril para tras.", ["Terra romeno com pausa"], ["Hipertrofia", "Forca"]],
    ["Afundo bulgaro com tronco inclinado", "Intermediario", "Composto", "Halteres e banco", ["Quadriceps", "Core"], "Inclina o tronco para aumentar a participacao de gluteos.", "Mantenha a canela dianteira mais vertical durante a descida.", ["Bulgaro com passada longa"], ["Hipertrofia"]],
    ["Frog pump", "Iniciante", "Isolado", "Peso corporal ou anilha", ["Adutores"], "Curta amplitude com alta sensacao de queima em gluteos.", "Una as plantas dos pes e contraia forte no topo.", ["Frog pump com banda"], ["Hipertrofia", "Resistencia"]],
    ["Cadeira abdutora inclinada", "Iniciante", "Isolado", "Maquina", ["Gluteo medio"], "Altera a inclinacao para maximizar gluteos laterais.", "Segure o tronco estavel e abra as pernas sem trancos.", ["Abducao sentada"], ["Hipertrofia"]],
    ["Pull-through na polia", "Iniciante", "Composto", "Polia e corda", ["Posterior de coxa", "Lombar"], "Padrao de dobradica com foco em extensao de quadril.", "A corda passa entre as pernas e o quadril guia o movimento.", ["Kettlebell swing tecnico"], ["Hipertrofia", "Resistencia"]],
    ["Good morning com barra leve", "Avancado", "Composto", "Barra", ["Posterior de coxa", "Lombar"], "Exercicio tecnico para cadeia posterior e controle lombopelvico.", "Use carga moderada e amplitude que preserve a coluna neutra.", ["Good morning sentado"], ["Forca", "Hipertrofia"]],
    ["Hip thrust no smith", "Intermediario", "Composto", "Smith machine e banco", ["Posterior de coxa", "Core"], "Versao guiada para gluteos com setup previsivel.", "Posicione a barra sobre o quadril e finalize com retroversao leve.", ["Hip thrust com banda"], ["Hipertrofia", "Forca"]],
  ]),
  Ombros: createExerciseCatalog([
    ["Desenvolvimento militar com barra", "Intermediario", "Composto", "Barra", ["Triceps", "Core"], "Press vertical classico para deltoides e forca geral.", "Gluteos e abdomen ativos para evitar hiperextensao lombar.", ["Push press", "Desenvolvimento sentado"], ["Forca", "Hipertrofia"], ["Desenvolvimento militar"]],
    ["Desenvolvimento com halteres sentado", "Iniciante", "Composto", "Halteres e banco", ["Triceps", "Core"], "Opcao estavel para ombros com amplitude segura.", "Mantenha punhos empilhados sobre os cotovelos.", ["Desenvolvimento neutro sentado"], ["Hipertrofia", "Iniciante"], ["Desenvolvimento com halteres"]],
    ["Arnold press", "Intermediario", "Composto", "Halteres", ["Triceps", "Deltoide anterior"], "Rotacao controlada para grande amplitude em deltoides.", "Nao acelere a rotacao e mantenha o tronco firme.", ["Arnold press sentado"], ["Hipertrofia"]],
    ["Desenvolvimento na maquina", "Iniciante", "Composto", "Maquina", ["Triceps"], "Variante guiada para acumular volume com menos demanda tecnica.", "Ajuste o banco para iniciar na altura do queixo.", ["Shoulder press convergente"], ["Hipertrofia", "Iniciante"]],
    ["Elevacao lateral com halteres", "Iniciante", "Isolado", "Halteres", ["Supraespinal"], "Base do treino para deltoide lateral.", "Suba ate a linha do ombro sem encolher o trapezio.", ["Elevacao lateral inclinada", "Elevacao lateral parcial"], ["Hipertrofia"], ["Elevacao lateral"]],
    ["Elevacao lateral na polia", "Intermediario", "Isolado", "Polia", ["Supraespinal"], "Mantem tensao constante desde o inicio do arco.", "Inicie de leve afastamento da polia para melhor linha de forca.", ["Elevacao lateral unilateral"], ["Hipertrofia"]],
    ["Elevacao frontal com halteres", "Iniciante", "Isolado", "Halteres", ["Peitoral superior"], "Foco em deltoide anterior com controle.", "Eleve sem usar embalo e desca lentamente.", ["Elevacao frontal alternada"], ["Hipertrofia"], ["Elevacao frontal"]],
    ["Elevacao frontal com barra", "Intermediario", "Isolado", "Barra", ["Peitoral superior"], "Permite carga maior para deltoide anterior.", "Segure a barra com pegada confortavel e sem jogar lombar.", ["Elevacao frontal com anilha"], ["Hipertrofia"]],
    ["Crucifixo inverso no peck deck", "Iniciante", "Isolado", "Maquina", ["Romboides", "Trapezio medio"], "Trabalha deltoide posterior com boa estabilidade.", "Peito apoiado e movimento abrindo pelos cotovelos.", ["Reverse fly na maquina"], ["Hipertrofia", "Postura"], ["Crucifixo inverso"]],
    ["Crucifixo inverso com halteres", "Intermediario", "Isolado", "Halteres", ["Romboides", "Trapezio medio"], "Variante livre para parte posterior dos ombros.", "Incline o tronco e mantenha a lombar neutra.", ["Reverse fly no banco inclinado"], ["Hipertrofia"]],
    ["Face pull na polia", "Iniciante", "Isolado", "Polia e corda", ["Rotadores externos", "Trapezio medio"], "Excelente para saude do ombro e deltoide posterior.", "Puxe a corda separando as pontas no final do movimento.", ["Face pull ajoelhado"], ["Reabilitacao", "Resistencia"], ["Face pull"]],
    ["Remada alta com barra W", "Intermediario", "Composto", "Barra W", ["Trapezio", "Biceps"], "Movimento para deltoide lateral e parte alta das costas.", "Suba os cotovelos acima das maos sem elevar demais a carga.", ["Remada alta na polia"], ["Hipertrofia"], ["Remada alta"]],
    ["Landmine press unilateral", "Intermediario", "Composto", "Barra landmine", ["Core", "Serratil"], "Press diagonal util para ombros e estabilidade.", "Mantenha costelas baixas e quadril alinhado.", ["Landmine press meio ajoelhado"], ["Forca", "Reabilitacao"]],
    ["Push press", "Avancado", "Composto", "Barra", ["Triceps", "Quadriceps"], "Usa impulso de pernas para sobrecarga acima da cabeca.", "Dip curto e explosivo, finalizando com tronco firme.", ["Push jerk"], ["Forca", "Potencia"]],
    ["Y-raise no banco inclinado", "Intermediario", "Isolado", "Halteres leves e banco", ["Trapezio inferior", "Rotadores externos"], "Fortalece deltoide e musculatura estabilizadora da escapula.", "Movimento controlado, polegares apontando para cima.", ["Y-raise na polia"], ["Reabilitacao", "Resistencia"]],
  ]),
  Biceps: createExerciseCatalog([
    ["Rosca direta com barra", "Iniciante", "Isolado", "Barra", ["Braquial", "Antebracos"], "Padrao classico para flexores do cotovelo.", "Cotovelos fixos ao lado do tronco e sem balanco.", ["Rosca direta no cabo"], ["Hipertrofia", "Forca"], ["Rosca direta"]],
    ["Rosca alternada com halteres", "Iniciante", "Isolado", "Halteres", ["Braquial", "Antebracos"], "Permite maior controle unilateral e supinacao.", "Supine progressivamente o punho durante a subida.", ["Rosca alternada sentado"], ["Hipertrofia"], ["Rosca alternada"]],
    ["Rosca martelo com halteres", "Iniciante", "Isolado", "Halteres", ["Braquial", "Braquiorradial"], "Fortalece biceps e antebraco com pegada neutra.", "Evite projetar os cotovelos para frente.", ["Rosca martelo alternada", "Rosca cross body"], ["Hipertrofia", "Forca"], ["Rosca martelo"]],
    ["Rosca Scott na maquina", "Iniciante", "Isolado", "Maquina", ["Braquial"], "Isolamento com estabilidade para a fase concentrica.", "Mantenha axilas bem apoiadas no banco Scott.", ["Rosca Scott com barra W"], ["Hipertrofia"], ["Rosca Scott"]],
    ["Rosca concentrada unilateral", "Intermediario", "Isolado", "Halter", ["Braquial"], "Excelente para foco em encurtamento do biceps.", "Apoie o cotovelo na parte interna da coxa e suba sem embalo.", ["Rosca concentrada sentada"], ["Hipertrofia"], ["Rosca concentrada"]],
    ["Rosca inversa com barra W", "Intermediario", "Isolado", "Barra W", ["Braquiorradial", "Extensores do antebraco"], "Fortalece antebracos e braquial com pegada pronada.", "Punhos neutros e cotovelos colados ao corpo.", ["Rosca inversa na polia"], ["Resistencia", "Hipertrofia"], ["Rosca inversa"]],
    ["Rosca na polia baixa", "Iniciante", "Isolado", "Polia", ["Braquial", "Antebracos"], "Tensao constante durante toda a flexao do cotovelo.", "Use postura ereta e evite balanço do tronco.", ["Rosca com corda", "Rosca com barra reta na polia"], ["Hipertrofia"]],
    ["Rosca spider com barra W", "Intermediario", "Isolado", "Barra W e banco inclinado", ["Braquial"], "Diminui roubos e aumenta foco no encurtamento.", "Peito apoiado no banco e descida completa com controle.", ["Spider curl com halteres"], ["Hipertrofia"]],
    ["Rosca 21 com barra", "Intermediario", "Isolado", "Barra", ["Braquial", "Antebracos"], "Metodo de series parciais e completas para alto estresse metabolico.", "Controle cada segmento da amplitude e nao use embalo.", ["Rosca 21 com halteres"], ["Hipertrofia", "Resistencia"]],
    ["Chin-up com foco em biceps", "Avancado", "Composto", "Peso corporal", ["Costas", "Braquial"], "Puxada supinada que combina dorsais e forte recrutamento de biceps.", "Inicie o movimento com peito alto e cotovelos para baixo.", ["Chin-up assistido"], ["Forca", "Hipertrofia"]],
    ["Rosca martelo na corda", "Iniciante", "Isolado", "Polia e corda", ["Braquial", "Braquiorradial"], "Mantem tensao continua em pegada neutra.", "Separe levemente as pontas da corda no final da subida.", ["Rosca corda alternada"], ["Hipertrofia"]],
    ["Rosca inclinada com halteres", "Intermediario", "Isolado", "Halteres e banco inclinado", ["Braquial"], "Alongamento forte do biceps pela posicao de ombro em extensao.", "Ombros relaxados e cotovelos apontando para baixo.", ["Rosca inclinada alternada"], ["Hipertrofia"]],
    ["Rosca preacher com halter", "Intermediario", "Isolado", "Halter e banco Scott", ["Braquial"], "Permite trabalhar unilateralmente em banco Scott.", "Nao perca a tensao no final da descida.", ["Preacher curl na polia"], ["Hipertrofia"]],
    ["Rosca cross body", "Iniciante", "Isolado", "Halteres", ["Braquial", "Braquiorradial"], "Martelo diagonal que melhora densidade do braco.", "Leve o halter em direcao ao ombro oposto sem girar o tronco.", ["Rosca martelo alternada"], ["Hipertrofia"]],
    ["Rosca Zottman", "Intermediario", "Isolado", "Halteres", ["Antebracos", "Braquial"], "Combina supinacao na subida e pronacao na descida.", "Gire os punhos com controle no topo da repeticao.", ["Rosca Zottman sentada"], ["Hipertrofia", "Resistencia"]],
  ]),
  Triceps: createExerciseCatalog([
    ["Triceps testa com barra W", "Intermediario", "Isolado", "Barra W e banco", ["Peitoral", "Deltoide anterior"], "Extensao de cotovelo com grande alongamento do triceps.", "Cotovelos apontados para cima e sem abrir excessivamente.", ["Skull crusher com halteres"], ["Hipertrofia", "Forca"], ["Triceps testa"]],
    ["Triceps frances com halter", "Iniciante", "Isolado", "Halter", ["Core"], "Extensao acima da cabeca para maior foco na cabeca longa.", "Mantenha cotovelos proximos e descida controlada atras da cabeca.", ["Frances bilateral", "Frances sentado"], ["Hipertrofia"], ["Triceps frances"]],
    ["Triceps corda na polia", "Iniciante", "Isolado", "Polia e corda", ["Antebracos"], "Variante confortavel e versatil para alto volume.", "Separe as pontas da corda no final da extensao.", ["Pushdown com corda unilateral"], ["Hipertrofia", "Resistencia"], ["Triceps corda"]],
    ["Triceps pulley com barra reta", "Iniciante", "Isolado", "Polia e barra reta", ["Antebracos"], "Pushdown tradicional para triceps com facil progressao.", "Cotovelos fixos junto ao corpo e punhos neutros.", ["Pushdown com barra V"], ["Hipertrofia"]],
    ["Mergulho entre bancos", "Iniciante", "Composto", "Peso corporal e bancos", ["Peitoral", "Deltoide anterior"], "Opcao acessivel de triceps com peso corporal.", "Desca ate confortavel para o ombro e suba sem tranco.", ["Mergulho com joelhos flexionados"], ["Resistencia", "Emagrecimento"], ["Mergulho"]],
    ["Supino fechado com barra", "Intermediario", "Composto", "Barra e banco", ["Peitoral", "Deltoide anterior"], "Press horizontal com maior enfase em triceps.", "Pegada fechada sem exagero e cotovelos apontando para baixo.", ["Supino fechado no smith"], ["Forca", "Hipertrofia"], ["Supino fechado"]],
    ["Triceps coice com halter", "Iniciante", "Isolado", "Halter", ["Deltoide posterior"], "Extensao de cotovelo em posicao inclinada.", "Fixe o braco paralelo ao tronco antes de estender.", ["Kickback na polia"], ["Hipertrofia"], ["Coice de triceps"]],
    ["Extensao acima da cabeca na polia", "Intermediario", "Isolado", "Polia e corda", ["Core"], "Excelente para cabeca longa do triceps.", "Incline levemente o tronco e mantenha os cotovelos altos.", ["Overhead extension unilateral"], ["Hipertrofia"]],
    ["Paralelas com foco em triceps", "Avancado", "Composto", "Barras paralelas", ["Peitoral", "Deltoide anterior"], "Movimento intenso com alta carga relativa para triceps.", "Tronco mais vertical e cotovelos apontando para tras.", ["Paralelas assistidas"], ["Forca", "Hipertrofia"]],
    ["Triceps unilateral na polia", "Iniciante", "Isolado", "Polia", ["Antebracos"], "Ajuda a equilibrar os dois lados com mais controle.", "Estabilize o ombro e estenda ate o final com controle.", ["Pushdown unilateral reverso"], ["Hipertrofia", "Reabilitacao"]],
    ["Triceps na maquina", "Iniciante", "Isolado", "Maquina", ["Antebracos"], "Versao guiada util para iniciantes e series de volume.", "Ajuste o assento para alinhar cotovelos ao eixo da maquina.", ["Dip machine"], ["Hipertrofia", "Iniciante"]],
    ["Triceps banco", "Iniciante", "Composto", "Peso corporal e banco", ["Peitoral", "Deltoide anterior"], "Movimento simples para triceps usando apoio no banco.", "Desca pouco se houver desconforto no ombro.", ["Bench dips com anilha"], ["Resistencia"], ["Triceps banco"]],
    ["JM press com barra", "Avancado", "Composto", "Barra e banco", ["Peitoral", "Deltoide anterior"], "Combina supino fechado e extensao para sobrecarregar triceps.", "Trajetoria curta com cotovelos controlados.", ["JM press no smith"], ["Forca", "Hipertrofia"]],
    ["Tate press", "Intermediario", "Isolado", "Halteres e banco", ["Peitoral"], "Variante de extensao para porcao medial do triceps.", "Mantenha cotovelos abertos com controle da carga.", ["Tate press alternado"], ["Hipertrofia"]],
    ["Flexao diamante", "Intermediario", "Composto", "Peso corporal", ["Peitoral", "Core"], "Flexao fechada para grande participacao de triceps.", "Mantenha maos proximas e tronco rigido do inicio ao fim.", ["Flexao diamante com joelhos apoiados"], ["Resistencia", "Emagrecimento"]],
  ]),
  Abdomen: createExerciseCatalog([
    ["Prancha frontal", "Iniciante", "Isometrico", "Peso corporal", ["Gluteos", "Ombros"], "Base de estabilidade anterior para o core.", "Contraia gluteos e mantenha lombar neutra sem afundar o quadril.", ["Prancha com apoio elevado", "Prancha com deslocamento"], ["Resistencia", "Reabilitacao"], ["Prancha"]],
    ["Prancha lateral", "Iniciante", "Isometrico", "Peso corporal", ["Gluteo medio", "Obliquos"], "Fortalece anti-flexao lateral e estabilidade do tronco.", "Empurre o antebraco no solo e alinhe ombro, quadril e tornozelos.", ["Prancha lateral com joelho apoiado"], ["Resistencia", "Reabilitacao"]],
    ["Crunch no solo", "Iniciante", "Isolado", "Peso corporal", ["Flexores de quadril"], "Flexao curta de tronco para reto abdominal.", "Tire as escapas do solo sem puxar o pescoco com as maos.", ["Crunch com pausa"], ["Hipertrofia", "Iniciante"], ["Crunch"]],
    ["Crunch na maquina", "Iniciante", "Isolado", "Maquina", ["Obliquos"], "Permite progressao de carga no reto abdominal.", "Expire no encurtamento e volte controlando a carga.", ["Crunch no cabo"], ["Hipertrofia"]],
    ["Abdominal infra no banco", "Iniciante", "Isolado", "Banco", ["Flexores de quadril"], "Eleva pelve e pernas com foco em porcao inferior do abdomen.", "Evite embalo das pernas e pense em enrolar a pelve.", ["Infra com joelhos flexionados"], ["Hipertrofia"], ["Abdominal infra"]],
    ["Elevacao de pernas na barra", "Avancado", "Composto", "Barra fixa", ["Flexores de quadril", "Obliquos"], "Desafia o core em suspensao com grande amplitude.", "Nao balance o corpo e inicie pela retroversao pelvica.", ["Elevacao de joelhos na barra"], ["Forca", "Hipertrofia"], ["Elevacao de pernas"]],
    ["Ab wheel", "Avancado", "Composto", "Roda abdominal", ["Latissimos", "Gluteos"], "Anti-extensao intensa para todo o core.", "Deslize apenas ate onde conseguir manter a lombar neutra.", ["Ab wheel ajoelhado", "Rollout na barra"], ["Forca", "Resistencia"]],
    ["Dead bug", "Iniciante", "Controle motor", "Peso corporal", ["Flexores profundos", "Gluteos"], "Padrao basico de estabilidade lombopelvica.", "Costas baixas coladas no solo durante o movimento alternado.", ["Dead bug com banda"], ["Reabilitacao", "Iniciante"]],
    ["Hollow hold", "Intermediario", "Isometrico", "Peso corporal", ["Flexores de quadril", "Serratil"], "Posicao global de anti-extensao muito usada em ginastica.", "Mantenha lombar colada no solo e costelas encaixadas.", ["Hollow rocks"], ["Resistencia", "Forca"]],
    ["Russian twist com anilha", "Intermediario", "Isolado", "Anilha", ["Obliquos", "Flexores de quadril"], "Rotacao controlada de tronco para obliquos.", "Gire o tronco, nao apenas os bracos, mantendo peito aberto.", ["Russian twist sem carga"], ["Resistencia", "Emagrecimento"]],
    ["Mountain climber", "Iniciante", "Composto", "Peso corporal", ["Ombros", "Flexores de quadril"], "Combina core e gasto energetico em alta cadencia.", "Quadris baixos e maos ativas empurrando o solo.", ["Mountain climber cruzado"], ["Emagrecimento", "Resistencia"]],
    ["Pallof press", "Iniciante", "Anti-rotacao", "Polia ou elastico", ["Obliquos", "Gluteos"], "Treina resistencia a rotacao para estabilidade funcional.", "Expulse o ar ao estender os bracos sem girar o tronco.", ["Pallof hold", "Pallof press ajoelhado"], ["Reabilitacao", "Resistencia"]],
    ["Abdominal na polia ajoelhado", "Intermediario", "Isolado", "Polia e corda", ["Flexores de quadril"], "Permite carga alta para flexao de tronco.", "Arredonde a coluna toracica, nao apenas puxe com os bracos.", ["Crunch na polia em pe"], ["Hipertrofia"], ["Abdominal na polia"]],
    ["Reverse crunch", "Iniciante", "Isolado", "Peso corporal", ["Flexores de quadril"], "Versao controlada de retroversao pelvica no solo.", "Suba o quadril suavemente sem impulsionar com as pernas.", ["Reverse crunch no banco"], ["Hipertrofia", "Iniciante"]],
    ["Sit-up com peso", "Intermediario", "Composto", "Anilha ou halter", ["Flexores de quadril"], "Variante mais dinamica para tronco completo.", "Controle a descida e evite perder a lombar no final.", ["Sit-up sem peso"], ["Resistencia", "Hipertrofia"]],
  ]),
  Panturrilhas: createExerciseCatalog([
    ["Panturrilha em pe na maquina", "Iniciante", "Isolado", "Maquina", ["Soleo"], "Trabalha gastrocnemio em grande amplitude.", "Desca ate alongar e suba ate o pico sem quicar.", ["Standing calf raise unilateral"], ["Hipertrofia", "Resistencia"], ["Panturrilha em pe"]],
    ["Panturrilha sentada na maquina", "Iniciante", "Isolado", "Maquina", ["Soleo"], "Enfatiza o soleo com joelhos flexionados.", "Mantenha quadril estavel e ritmo controlado.", ["Panturrilha sentada unilateral"], ["Hipertrofia"], ["Panturrilha sentado"]],
    ["Panturrilha no leg press", "Iniciante", "Isolado", "Leg press", ["Soleo"], "Opcao segura para sobrecarga progressiva.", "Movimente apenas os tornozelos sem dobrar os joelhos.", ["Panturrilha no leg press unilateral"], ["Hipertrofia"], ["Panturrilha no leg press"]],
    ["Panturrilha unilateral em degrau", "Iniciante", "Isolado", "Peso corporal e degrau", ["Soleo"], "Ajuda a corrigir assimetrias e melhorar amplitude.", "Segure apoio leve e use o calcanhar para alongar bem.", ["Panturrilha unilateral com halter"], ["Resistencia", "Reabilitacao"], ["Panturrilha unilateral"]],
    ["Donkey calf raise", "Intermediario", "Isolado", "Maquina ou apoio", ["Soleo"], "Variante classica com grande alongamento em dorsiflexao.", "Quadril dobrado e coluna neutra durante toda a serie.", ["Donkey calf machine"], ["Hipertrofia"], ["Donkey calf raise"]],
    ["Panturrilha no smith em pe", "Intermediario", "Isolado", "Smith machine", ["Soleo"], "Permite progressao consistente em pe com estabilidade.", "Use plataforma sob a ponta dos pes para ganhar amplitude.", ["Panturrilha no smith unilateral"], ["Hipertrofia", "Forca"]],
    ["Panturrilha no hack machine", "Intermediario", "Isolado", "Hack machine", ["Soleo"], "Boa alternativa para sobrecarga alta em amplitude curta.", "Segure a plataforma firme e nao trave joelhos em excesso.", ["Hack calf raise reverso"], ["Hipertrofia"]],
    ["Saltos na ponta dos pes", "Iniciante", "Pliometrico", "Peso corporal", ["Quadriceps", "Gluteos"], "Trabalha reatividade e resistencia de panturrilhas.", "Aterrisse macio e mantenha tornozelos ativos.", ["Pogo jumps"], ["Resistencia", "Condicionamento"]],
    ["Farmer walk na ponta dos pes", "Intermediario", "Funcional", "Halteres", ["Core", "Antebracos"], "Combina isometria de panturrilha com estabilidade global.", "Passos curtos e controlados mantendo o calcanhar elevado.", ["Suitcase walk na ponta dos pes"], ["Resistencia", "Emagrecimento"]],
    ["Panturrilha isometrica em degrau", "Iniciante", "Isometrico", "Peso corporal e degrau", ["Soleo"], "Mantem alta tensao sustentada no pico de contracao.", "Segure o topo sem perder alinhamento dos tornozelos.", ["Isometria unilateral em degrau"], ["Reabilitacao", "Resistencia"]],
    ["Panturrilha com joelhos flexionados em pe", "Intermediario", "Isolado", "Smith ou halteres", ["Soleo"], "Variante em pe com foco maior em soleo.", "Mantenha joelhos semiflexionados durante toda a amplitude.", ["Bent-knee calf raise"], ["Hipertrofia"]],
    ["Panturrilha na maquina horizontal", "Iniciante", "Isolado", "Maquina", ["Soleo"], "Opcao guiada para volume com estabilidade.", "Ritmo controlado e pause um segundo no pico da subida.", ["Panturrilha horizontal unilateral"], ["Hipertrofia"]],
    ["Pogo jumps", "Intermediario", "Pliometrico", "Peso corporal", ["Quadriceps", "Tibial anterior"], "Saltos curtos e rapidos para elasticidade do tornozelo.", "Mantenha contato minimo com o solo e tronco alto.", ["Line hops"], ["Condicionamento", "Resistencia"]],
    ["Panturrilha unilateral sentada", "Intermediario", "Isolado", "Maquina ou anilha", ["Soleo"], "Permite ajustar melhor carga lado a lado.", "Mantenha joelho estavel e amplitude completa.", ["Panturrilha sentada bilateral"], ["Hipertrofia"]],
    ["Panturrilha no step com halteres", "Iniciante", "Isolado", "Halteres e step", ["Soleo"], "Versao simples para treinar em casa ou academia.", "Controle a descida abaixo do nivel do step antes de subir.", ["Panturrilha bilateral no step"], ["Resistencia", "Hipertrofia"]],
  ]),
});

const CATEGORY_PRESENTATION = Object.freeze({
  Peito: { subtitle: "Supinos, crucifixos e variacoes de empurrar." },
  Costas: { subtitle: "Puxadas, remadas e movimentos de dorsal." },
  Pernas: { subtitle: "Base, quadriceps e movimentos principais." },
  Gluteos: { subtitle: "Extensao de quadril e foco em cadeia posterior." },
  Ombros: { subtitle: "Presses, elevacoes e estabilidade escapular." },
  Biceps: { subtitle: "Roscas e variacoes de flexao de cotovelo." },
  Triceps: { subtitle: "Extensoes, empurradas e finalizacoes." },
  Abdomen: { subtitle: "Core, controle postural e estabilizacao." },
  Panturrilhas: { subtitle: "Flexao plantar, volume e resistencia local." },
});

const SPLIT_PRESET_LIBRARY = Object.freeze({
  Push: {
    groups: ["Peito", "Ombros", "Triceps"],
    exercises: [
      { category: "Peito", name: "Supino reto com barra", sets: "4", reps: "8", load: "" },
      { category: "Peito", name: "Crossover na polia alta", sets: "3", reps: "12", load: "" },
      { category: "Ombros", name: "Desenvolvimento com halteres sentado", sets: "4", reps: "10", load: "" },
      { category: "Ombros", name: "Elevacao lateral com halteres", sets: "3", reps: "12", load: "" },
      { category: "Triceps", name: "Triceps corda na polia", sets: "3", reps: "12", load: "" },
    ],
  },
  Pull: {
    groups: ["Costas", "Biceps"],
    exercises: [
      { category: "Costas", name: "Puxada frontal pronada", sets: "4", reps: "10", load: "" },
      { category: "Costas", name: "Remada baixa na polia", sets: "4", reps: "10", load: "" },
      { category: "Costas", name: "Barra fixa pronada", sets: "3", reps: "8", load: "" },
      { category: "Biceps", name: "Rosca direta com barra", sets: "3", reps: "10", load: "" },
      { category: "Biceps", name: "Rosca martelo com halteres", sets: "3", reps: "12", load: "" },
    ],
  },
  Legs: {
    groups: ["Pernas", "Gluteos", "Panturrilhas"],
    exercises: [
      { category: "Pernas", name: "Agachamento livre com barra", sets: "4", reps: "8", load: "" },
      { category: "Pernas", name: "Leg press 45", sets: "4", reps: "12", load: "" },
      { category: "Pernas", name: "Mesa flexora", sets: "3", reps: "12", load: "" },
      { category: "Gluteos", name: "Elevacao pelvica com barra", sets: "4", reps: "10", load: "" },
      { category: "Panturrilhas", name: "Panturrilha em pe na maquina", sets: "4", reps: "15", load: "" },
    ],
  },
  Upper: {
    groups: ["Peito", "Costas", "Ombros", "Biceps", "Triceps"],
    exercises: [
      { category: "Peito", name: "Supino inclinado com barra", sets: "4", reps: "8", load: "" },
      { category: "Costas", name: "Remada curvada com barra", sets: "4", reps: "8", load: "" },
      { category: "Ombros", name: "Arnold press", sets: "3", reps: "10", load: "" },
      { category: "Biceps", name: "Rosca alternada com halteres", sets: "3", reps: "10", load: "" },
      { category: "Triceps", name: "Triceps testa com barra W", sets: "3", reps: "10", load: "" },
    ],
  },
  Lower: {
    groups: ["Pernas", "Gluteos", "Panturrilhas", "Abdomen"],
    exercises: [
      { category: "Pernas", name: "Hack squat", sets: "4", reps: "10", load: "" },
      { category: "Pernas", name: "Stiff com barra", sets: "4", reps: "10", load: "" },
      { category: "Gluteos", name: "Coice no cabo", sets: "3", reps: "12", load: "" },
      { category: "Panturrilhas", name: "Panturrilha sentada na maquina", sets: "4", reps: "15", load: "" },
      { category: "Abdomen", name: "Prancha frontal", sets: "3", reps: "30", load: "" },
    ],
  },
  Core: {
    groups: ["Abdomen", "Ombros"],
    exercises: [
      { category: "Abdomen", name: "Crunch no solo", sets: "3", reps: "15", load: "" },
      { category: "Abdomen", name: "Elevacao de pernas na barra", sets: "3", reps: "12", load: "" },
      { category: "Abdomen", name: "Ab wheel", sets: "3", reps: "10", load: "" },
      { category: "Ombros", name: "Face pull na polia", sets: "3", reps: "15", load: "" },
      { category: "Ombros", name: "Crucifixo inverso no peck deck", sets: "3", reps: "12", load: "" },
    ],
  },
});

const ALL_EXERCISE_DETAILS = Object.freeze(
  Object.entries(EXERCISE_LIBRARY).flatMap(([category, exercises]) =>
    exercises.map((exercise) => ({ ...exercise, category })),
  ),
);
const ALL_EXERCISES = ALL_EXERCISE_DETAILS.map((exercise) => exercise.name);
const EXERCISE_DETAILS_BY_NAME = new Map(
  ALL_EXERCISE_DETAILS.flatMap((exercise) => [
    [normalizeLookupValue(exercise.name), exercise],
    ...exercise.aliases.map((alias) => [normalizeLookupValue(alias), exercise]),
  ]),
);
const EXERCISE_CATEGORY_BY_NAME = new Map(
  Array.from(EXERCISE_DETAILS_BY_NAME.entries(), ([normalizedName, exercise]) => [normalizedName, exercise.category]),
);

const elements = {
  trainingName: document.querySelector("#training-name"),
  addTrainingButton: document.querySelector("#add-training-button"),
  downloadButton: document.querySelector("#download-button"),
  downloadAllButton: document.querySelector("#download-all-button"),
  removeTrainingButton: document.querySelector("#remove-training-button"),
  trainingTabs: document.querySelector("#training-tabs"),
  addExerciseButton: document.querySelector("#add-exercise-button"),
  clearTrainingButton: document.querySelector("#clear-training-button"),
  tableBody: document.querySelector("#exercise-table-body"),
  emptyState: document.querySelector("#empty-state"),
  rowTemplate: document.querySelector("#exercise-row-template"),
  activeTrainingLabel: document.querySelector("#active-training-label"),
  exerciseCount: document.querySelector("#exercise-count"),
  seriesCount: document.querySelector("#series-count"),
  lastSavedLabel: document.querySelector("#last-saved-label"),
  statusMessage: document.querySelector("#status-message"),
  activeSplitLabel: document.querySelector("#active-split-label"),
  splitPresets: document.querySelector("#split-presets"),
  briefingBoard: document.querySelector("#briefing-board"),
  traineeName: document.querySelector("#trainee-name"),
  trainingGoal: document.querySelector("#training-goal"),
  trainingLevel: document.querySelector("#training-level"),
  trainingDate: document.querySelector("#training-date"),
  sessionDuration: document.querySelector("#session-duration"),
  restTime: document.querySelector("#rest-time"),
  coachName: document.querySelector("#coach-name"),
  weeklyFrequency: document.querySelector("#weekly-frequency"),
  briefingNotes: document.querySelector("#briefing-notes"),
  briefingChipDate: document.querySelector("#briefing-chip-date"),
  briefingChipGoal: document.querySelector("#briefing-chip-goal"),
  briefingChipLevel: document.querySelector("#briefing-chip-level"),
  briefingChipFrequency: document.querySelector("#briefing-chip-frequency"),
  toolbarTrainingName: document.querySelector("#toolbar-training-name"),
  toolbarExerciseCount: document.querySelector("#toolbar-exercise-count"),
  exerciseDrawer: document.querySelector("#exercise-drawer"),
  exerciseDrawerTraining: document.querySelector("#exercise-drawer-training"),
  exerciseDrawerCategory: document.querySelector("#exercise-drawer-category"),
  exerciseDrawerCategoryBadge: document.querySelector("#exercise-drawer-category-badge"),
  exerciseDrawerDescription: document.querySelector("#exercise-drawer-description"),
  exerciseDrawerSearch: document.querySelector("#exercise-drawer-search"),
  exerciseDrawerMeta: document.querySelector("#exercise-drawer-meta"),
  exerciseDrawerMetaCopy: document.querySelector("#exercise-drawer-meta-copy"),
  exerciseDrawerList: document.querySelector("#exercise-drawer-list"),
  categoryDrawer: document.querySelector("#category-drawer"),
  categoryDrawerTraining: document.querySelector("#category-drawer-training"),
  categoryDrawerCurrent: document.querySelector("#category-drawer-current"),
  categoryDrawerDescription: document.querySelector("#category-drawer-description"),
  categoryDrawerList: document.querySelector("#category-drawer-list"),
};

const exerciseDrawerState = {
  exerciseId: null,
  query: "",
  previewName: "",
};

const categoryDrawerState = {
  exerciseId: null,
};

const briefingFieldMap = {
  traineeName: "traineeName",
  trainingGoal: "goal",
  trainingLevel: "level",
  trainingDate: "sessionDate",
  sessionDuration: "duration",
  restTime: "rest",
  coachName: "coachName",
  weeklyFrequency: "frequency",
  briefingNotes: "notes",
};

const defaultExercise = () => ({
  id: crypto.randomUUID(),
  category: "",
  name: "",
  sets: "",
  reps: "",
  load: "",
  notes: "",
});

const createEmptyProfile = () => ({
  traineeName: "",
  goal: "",
  level: "",
  sessionDate: "",
  duration: "",
  rest: "",
  coachName: "",
  frequency: "",
  notes: "",
});

const defaultState = () => ({
  activeTrainingId: "training-a",
  lastSavedAt: new Date().toISOString(),
  profile: createEmptyProfile(),
  trainings: [
    { id: "training-a", name: "Treino A", splitLabel: "", focusGroups: [], exercises: [defaultExercise()] },
    { id: "training-b", name: "Treino B", splitLabel: "", focusGroups: [], exercises: [] },
    { id: "training-c", name: "Treino C", splitLabel: "", focusGroups: [], exercises: [] },
    { id: "training-d", name: "Treino D", splitLabel: "", focusGroups: [], exercises: [] },
  ],
});

let state = loadState();

updatePrintMetadata();
render();
bindEvents();

function bindEvents() {
  elements.addTrainingButton?.addEventListener("click", handleAddTraining);
  elements.addExerciseButton?.addEventListener("click", handleAddExercise);
  elements.clearTrainingButton?.addEventListener("click", handleClearTraining);
  elements.downloadButton?.addEventListener("click", handleDownload);
  elements.downloadAllButton?.addEventListener("click", handleDownloadAll);
  elements.removeTrainingButton?.addEventListener("click", handleRemoveTraining);
  elements.trainingName?.addEventListener("input", handleTrainingRename);
  Object.keys(briefingFieldMap).forEach((elementKey) => {
    elements[elementKey]?.addEventListener("input", handleBriefingFieldUpdate);
    elements[elementKey]?.addEventListener("change", handleBriefingFieldUpdate);
  });
  elements.tableBody.addEventListener("focusin", handleTableFocusIn);
  elements.tableBody.addEventListener("input", handleTableFieldUpdate);
  elements.tableBody.addEventListener("change", handleTableFieldUpdate);
  elements.tableBody.addEventListener("keydown", handleTableKeyDown);
  elements.tableBody.addEventListener("click", handleTableClick);
  elements.tableBody.addEventListener("dragstart", handleTableDragStart);
  elements.tableBody.addEventListener("dragover", handleTableDragOver);
  elements.tableBody.addEventListener("drop", handleTableDrop);
  elements.tableBody.addEventListener("dragend", handleTableDragEnd);
  elements.trainingTabs.addEventListener("click", handleTabClick);
  elements.exerciseDrawerSearch?.addEventListener("input", handleExerciseDrawerSearch);
  elements.exerciseDrawerList?.addEventListener("mouseover", handleExerciseDrawerPreview);
  elements.exerciseDrawerList?.addEventListener("focusin", handleExerciseDrawerPreview);
  document.addEventListener("click", handleDocumentClick);
  document.addEventListener("keydown", handleDocumentKeyDown);
}

function loadState() {
  try {
    const rawState = LEGACY_STORAGE_KEYS
      .map((storageKey) => localStorage.getItem(storageKey))
      .find(Boolean);

    if (!rawState) {
      return defaultState();
    }

    const parsedState = JSON.parse(rawState);
    const trainings = Array.isArray(parsedState.trainings) ? parsedState.trainings : [];

    if (!trainings.length) {
      return defaultState();
    }

    return {
      activeTrainingId: parsedState.activeTrainingId || trainings[0].id,
      lastSavedAt: parsedState.lastSavedAt || new Date().toISOString(),
      profile: {
        ...createEmptyProfile(),
        ...(parsedState.profile || trainings.find((training) => training.profile)?.profile || {}),
      },
      trainings: trainings.map((training, trainingIndex) => ({
        id: training.id || `training-${trainingIndex + 1}`,
        name: training.name || `Treino ${trainingIndex + 1}`,
        splitLabel: training.splitLabel || "",
        focusGroups: Array.isArray(training.focusGroups) ? training.focusGroups : [],
        exercises: Array.isArray(training.exercises)
          ? training.exercises.map((exercise) => ({
              ...defaultExercise(),
              ...exercise,
              id: exercise.id || crypto.randomUUID(),
              category: exercise.category || inferCategoryFromExerciseName(exercise.name) || "",
            }))
          : [],
      })),
    };
  } catch {
    return defaultState();
  }
}

function saveState() {
  state.lastSavedAt = new Date().toISOString();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  elements.lastSavedLabel.textContent = formatTimestamp(state.lastSavedAt);
}

function getActiveTraining() {
  const activeTraining = state.trainings.find((training) => training.id === state.activeTrainingId);

  if (activeTraining) {
    return activeTraining;
  }

  state.activeTrainingId = state.trainings[0]?.id || defaultState().activeTrainingId;
  return state.trainings[0] || defaultState().trainings[0];
}

function render() {
  syncExerciseDrawerState();
  renderTabs();
  renderTrainingHeader();
  renderPlanningBoard();
  renderBriefingBoard();
  renderTable();
  renderStats();
}

function renderPlanningBoard() {
  const activeTraining = getActiveTraining();

  elements.activeSplitLabel.textContent = activeTraining.splitLabel
    ? `${activeTraining.splitLabel} • ${activeTraining.focusGroups.join(", ")}`
    : "Sem preset aplicado.";

  elements.splitPresets.querySelectorAll("[data-action='apply-split']").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.split === activeTraining.splitLabel);
  });
}

function renderTabs() {
  elements.trainingTabs.innerHTML = "";

  state.trainings.forEach((training) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `tab-button${training.id === state.activeTrainingId ? " is-active" : ""}`;
    button.dataset.trainingId = training.id;
    button.textContent = getDisplayTrainingName(training.name);
    elements.trainingTabs.appendChild(button);
  });
}

function renderTrainingHeader() {
  const activeTraining = getActiveTraining();
  elements.trainingName.value = activeTraining.name;
  elements.trainingName.classList.toggle("is-invalid", !activeTraining.name.trim());
  elements.activeTrainingLabel.textContent = getDisplayTrainingName(activeTraining.name);
  if (elements.toolbarTrainingName) {
    elements.toolbarTrainingName.textContent = getDisplayTrainingName(activeTraining.name);
  }

  if (elements.removeTrainingButton instanceof HTMLButtonElement) {
    elements.removeTrainingButton.disabled = state.trainings.length === 1;
  }

  if (elements.downloadAllButton instanceof HTMLButtonElement) {
    elements.downloadAllButton.disabled = state.trainings.length <= 1;
  }
}

function renderBriefingBoard() {
  const profile = {
    ...createEmptyProfile(),
    ...(state.profile || {}),
  };

  Object.entries(briefingFieldMap).forEach(([elementKey, profileKey]) => {
    const field = elements[elementKey];
    if (field) {
      field.value = profile[profileKey] || "";
    }
  });

  elements.briefingChipDate.textContent = profile.sessionDate || "Nao definida";
  elements.briefingChipGoal.textContent = profile.goal || "Livre";
  elements.briefingChipLevel.textContent = profile.level || "Nao informado";
  elements.briefingChipFrequency.textContent = profile.frequency ? `${profile.frequency}x semana` : "0x semana";
}

function renderTable() {
  const activeTraining = getActiveTraining();
  elements.tableBody.innerHTML = "";

  activeTraining.exercises.forEach((exercise, index) => {
    const row = elements.rowTemplate.content.firstElementChild.cloneNode(true);
    const nameInput = row.querySelector('[data-field="name"]');

    row.dataset.exerciseId = exercise.id;
    row.querySelector(".sheet-table__index").textContent = String(index + 1).padStart(2, "0");

    renderCategoryPicker(row, exercise.category);

    row.querySelectorAll("[data-field]").forEach((fieldElement) => {
      const field = fieldElement.dataset.field;
      fieldElement.value = exercise[field] ?? "";

      if (field !== "category") {
        validateInputField(fieldElement, field, exercise[field]);
      }
    });

    nameInput.setAttribute("spellcheck", "false");
    elements.tableBody.appendChild(row);
  });

  elements.emptyState.hidden = activeTraining.exercises.length > 0;
}

function renderStats() {
  const activeTraining = getActiveTraining();
  const totalExercises = activeTraining.exercises.filter((exercise) => isExerciseMeaningful(exercise)).length;
  const totalSets = activeTraining.exercises.reduce((sum, exercise) => {
    const numericValue = Number(exercise.sets);
    return Number.isFinite(numericValue) ? sum + numericValue : sum;
  }, 0);

  elements.exerciseCount.textContent = String(totalExercises);
  elements.seriesCount.textContent = String(totalSets);
  elements.lastSavedLabel.textContent = formatTimestamp(state.lastSavedAt);

  if (elements.toolbarExerciseCount) {
    elements.toolbarExerciseCount.textContent = `${totalExercises} exercicio${totalExercises === 1 ? "" : "s"}`;
  }
}

function handleAddTraining() {
  const nextLabel = getNextTrainingName();
  const nextTraining = {
    id: crypto.randomUUID(),
    name: nextLabel,
    splitLabel: "",
    focusGroups: [],
    exercises: [],
  };

  state.trainings.push(nextTraining);
  state.activeTrainingId = nextTraining.id;
  persistAndRender();
  setStatusMessage("Novo treino criado.", "success");
}

function handleRemoveTraining() {
  const activeTraining = getActiveTraining();

  if (state.trainings.length === 1) {
    const confirmedReset = window.confirm(
      `Este e o ultimo treino. Deseja limpar ${getDisplayTrainingName(activeTraining.name)} e manter a estrutura base?`,
    );

    if (!confirmedReset) {
      return;
    }

    activeTraining.name = "Treino A";
    activeTraining.splitLabel = "";
    activeTraining.focusGroups = [];
    activeTraining.exercises = [];
    state.profile = createEmptyProfile();
    persistAndRender();
    setStatusMessage("Treino base redefinido.", "success");
    return;
  }

  const confirmed = window.confirm(`Remover ${getDisplayTrainingName(activeTraining.name)}? Esta acao nao pode ser desfeita.`);

  if (!confirmed) {
    return;
  }

  const activeTrainingIndex = state.trainings.findIndex((training) => training.id === activeTraining.id);
  state.trainings = state.trainings.filter((training) => training.id !== activeTraining.id);
  const fallbackTraining = state.trainings[Math.max(0, activeTrainingIndex - 1)] || state.trainings[0];
  state.activeTrainingId = fallbackTraining?.id || defaultState().activeTrainingId;
  persistAndRender();
  setStatusMessage("Treino removido.", "success");
}

function handleAddExercise() {
  const activeTraining = getActiveTraining();
  activeTraining.exercises.push(defaultExercise());
  persistAndRender();
  setStatusMessage("Nova linha adicionada. Escolha um agrupamento muscular para ver sugestoes.", "success");
}

function handleClearTraining() {
  const activeTraining = getActiveTraining();

  if (!activeTraining.exercises.length) {
    return;
  }

  const confirmed = window.confirm(`Limpar todos os exercicios de ${getDisplayTrainingName(activeTraining.name)}?`);

  if (!confirmed) {
    return;
  }

  activeTraining.exercises = [];
  persistAndRender();
  setStatusMessage("Treino limpo.", "success");
}

function handleDownload() {
  const exportPayload = createExportPayload(false);
  handlePdfDownload(exportPayload, false);
}

function handleDownloadAll() {
  if (state.trainings.length <= 1) {
    handleDownload();
    return;
  }

  const exportPayload = createExportPayload(true);
  handlePdfDownload(exportPayload, true);
}

function handlePdfDownload(payload, includeAllTrainings = true) {
  try {
    gerarPdf(payload, includeAllTrainings ? "treino.pdf" : "treino.pdf");
    setStatusMessage(
      includeAllTrainings ? "Todos os treinos exportados em PDF." : "Treino ativo exportado em PDF.",
      "success",
    );
  } catch (error) {
    setStatusMessage(error.message || "Falha ao gerar o PDF.", "error");
  }
}

function handlePrint() {
  updatePrintMetadata();
  document.body.classList.add("print-mode");
  window.print();
  window.setTimeout(() => {
    document.body.classList.remove("print-mode");
  }, 150);
}

function handleImportButtonClick() {
  elements.importFileInput.click();
}

async function handleImportFileChange(event) {
  const [file] = Array.from(event.target.files || []);

  if (!file) {
    return;
  }

  try {
    const importedTrainings = await parseImportFile(file);

    if (!importedTrainings.length) {
      throw new Error("Nenhum treino valido foi encontrado no arquivo.");
    }

    state.trainings = importedTrainings;
    state.activeTrainingId = importedTrainings[0].id;
    persistAndRender();
    setStatusMessage(`Importacao concluida: ${importedTrainings.length} treino(s) carregado(s).`, "success");
  } catch (error) {
    setStatusMessage(error.message || "Falha ao importar o arquivo.", "error");
  } finally {
    event.target.value = "";
  }
}

function handleTrainingRename(event) {
  const activeTraining = getActiveTraining();
  activeTraining.name = event.target.value.slice(0, 30);
  elements.trainingName.classList.toggle("is-invalid", !activeTraining.name.trim());
  elements.activeTrainingLabel.textContent = getDisplayTrainingName(activeTraining.name);
  renderTabs();
  persistAndRender(false);
}

function handleBriefingFieldUpdate(event) {
  const target = event.target;

  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement)) {
    return;
  }

  const matchedEntry = Object.entries(briefingFieldMap).find(([elementKey]) => elements[elementKey] === target);
  if (!matchedEntry) {
    return;
  }

  state.profile[matchedEntry[1]] = target.value;
  renderBriefingBoard();
  updatePrintMetadata();
  persistAndRender(false);
}

function handleTableFocusIn(event) {
  const target = event.target;

  if (!(target instanceof HTMLInputElement) || target.dataset.field !== "name") {
    return;
  }

  const row = target.closest("tr");
  const exercise = getExerciseById(row?.dataset.exerciseId);

  if (!exercise || !row) {
    return;
  }

  openExerciseDrawer(row, exercise);
}

function handleTableFieldUpdate(event) {
  const target = event.target;

  if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement) || !target.dataset.field) {
    return;
  }

  const row = target.closest("tr");
  const exercise = getExerciseById(row?.dataset.exerciseId);

  if (!exercise) {
    return;
  }

  const field = target.dataset.field;

  if (field === "category") {
    return;
  }

  exercise[field] = sanitizeValue(field, target.value);

  if (field === "name") {
    return;
  }

  validateInputField(target, field, exercise[field]);
  persistAndRender(false);
  renderStats();
}

function handleTableKeyDown(event) {
  const row = event.target.closest("tr");
  if (!row) {
    return;
  }

  if (event.key === "Escape") {
    closeExerciseDrawer();
    closeCategoryDrawer();
    return;
  }

  const target = event.target;
  const exercise = getExerciseById(row.dataset.exerciseId);
  const isNameField = target instanceof HTMLInputElement && target.dataset.field === "name";

  if (!exercise || !isNameField) {
    return;
  }

  if (["Enter", "ArrowDown", " "].includes(event.key)) {
    event.preventDefault();
    openExerciseDrawer(row, exercise);
  }
}

function handleTableClick(event) {
  const target = event.target;

  if (!(target instanceof HTMLElement)) {
    return;
  }

  const removeButton = target.closest('[data-action="remove"]');
  if (removeButton) {
    const row = removeButton.closest("tr");
    const activeTraining = getActiveTraining();
    activeTraining.exercises = activeTraining.exercises.filter(
      (exercise) => exercise.id !== row?.dataset.exerciseId,
    );
    persistAndRender();
    setStatusMessage("Exercicio removido.", "success");
    return;
  }

  const duplicateButton = target.closest('[data-action="duplicate"]');
  if (duplicateButton) {
    const row = duplicateButton.closest("tr");
    const exercise = getExerciseById(row?.dataset.exerciseId);

    if (!row || !exercise) {
      return;
    }

    duplicateExercise(exercise.id);
    return;
  }

  const toggleCategoryButton = target.closest('[data-action="toggle-category"]');
  if (toggleCategoryButton) {
    const row = toggleCategoryButton.closest("tr");
    const exercise = getExerciseById(row?.dataset.exerciseId);

    if (!row || !exercise) {
      return;
    }

    openCategoryDrawer(row, exercise);
    return;
  }

  const categoryOption = target.closest('[data-action="choose-category-drawer"]');
  if (categoryOption) {
    selectCategoryFromDrawer(categoryOption);
    return;
  }

  const toggleButton = target.closest('[data-action="toggle-suggestions"]');
  if (toggleButton) {
    const row = toggleButton.closest("tr");
    const exercise = getExerciseById(row?.dataset.exerciseId);

    if (!row || !exercise) {
      return;
    }

    openExerciseDrawer(row, exercise);
    return;
  }
}

function handleTableDragStart(event) {
  const target = event.target;

  if (!(target instanceof HTMLElement) || !target.closest('[data-action="drag-row"]')) {
    return;
  }

  const row = target.closest("tr");
  if (!row) {
    return;
  }

  row.classList.add("is-dragging");
  event.dataTransfer?.setData("text/plain", row.dataset.exerciseId || "");
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
  }
}

function handleTableDragOver(event) {
  const row = event.target instanceof HTMLElement ? event.target.closest("tr") : null;

  if (!row) {
    return;
  }

  event.preventDefault();
  elements.tableBody.querySelectorAll("tr").forEach((tableRow) => tableRow.classList.remove("is-drop-target"));
  row.classList.add("is-drop-target");
}

function handleTableDrop(event) {
  const targetRow = event.target instanceof HTMLElement ? event.target.closest("tr") : null;
  const sourceId = event.dataTransfer?.getData("text/plain");

  if (!targetRow || !sourceId || sourceId === targetRow.dataset.exerciseId) {
    return;
  }

  event.preventDefault();
  reorderExercise(sourceId, targetRow.dataset.exerciseId || "");
}

function handleTableDragEnd() {
  elements.tableBody.querySelectorAll("tr").forEach((row) => {
    row.classList.remove("is-dragging", "is-drop-target");
  });
}

function handleDocumentClick(event) {
  const target = event.target;

  if (target instanceof Element) {
    const closeCategoryDrawerButton = target.closest('[data-action="close-category-drawer"]');
    if (closeCategoryDrawerButton) {
      closeCategoryDrawer();
      return;
    }

    const closeDrawerButton = target.closest('[data-action="close-exercise-drawer"]');
    if (closeDrawerButton) {
      closeExerciseDrawer();
      return;
    }

    const drawerOption = target.closest('[data-action="choose-exercise-drawer"]');
    if (drawerOption) {
      selectExerciseFromDrawer(drawerOption);
      return;
    }

    const categoryOption = target.closest('[data-action="choose-category-drawer"]');
    if (categoryOption) {
      selectCategoryFromDrawer(categoryOption);
      return;
    }

    const splitButton = target.closest('[data-action="apply-split"]');
    if (splitButton) {
      applySplitPreset(splitButton.dataset.split || "", (splitButton.dataset.groups || "").split(",").filter(Boolean));
      return;
    }

  }

  if (target instanceof Element && (target.closest(".exercise-picker") || target.closest(".category-picker") || target.closest(".exercise-drawer__panel") || target.closest(".category-drawer__panel"))) {
    return;
  }
}

function handleDocumentKeyDown(event) {
  if (event.key === "Escape") {
    closeExerciseDrawer();
    closeCategoryDrawer();
  }
}

function handleTabClick(event) {
  const target = event.target;

  if (!(target instanceof HTMLButtonElement) || !target.dataset.trainingId) {
    return;
  }

  state.activeTrainingId = target.dataset.trainingId;
  persistAndRender();
}

function getExerciseById(exerciseId) {
  if (!exerciseId) {
    return null;
  }

  return getActiveTraining().exercises.find((exercise) => exercise.id === exerciseId) ?? null;
}

function sanitizeValue(field, value) {
  if (["sets", "reps", "load"].includes(field)) {
    return value.replace(/[^\d.,-]/g, "").replace(",", ".");
  }

  return value;
}

function validateInputField(input, field, value) {
  let isInvalid = false;

  if (["sets", "reps"].includes(field) && String(value).trim()) {
    isInvalid = Number(value) <= 0;
  }

  if (field === "load" && String(value).trim()) {
    isInvalid = Number(value) < 0;
  }

  input.classList.toggle("is-invalid", isInvalid);
}

function renderCategoryPicker(row, selectedCategory) {
  const hiddenInput = row.querySelector('[data-field="category"]');
  const trigger = row.querySelector('[data-action="toggle-category"]');

  if (!(hiddenInput instanceof HTMLInputElement) || !(trigger instanceof HTMLButtonElement)) {
    return;
  }

  hiddenInput.value = selectedCategory || "";
  trigger.textContent = selectedCategory || "Agrupamento";
  trigger.classList.toggle("is-placeholder", !selectedCategory);
}

function setExerciseCategory(row, exercise, category, openSuggestions = true) {
  const previousCategory = exercise.category;
  exercise.category = category || "";

  if (previousCategory !== exercise.category) {
    const inferredCategory = inferCategoryFromExerciseName(exercise.name);

    if (!exercise.category || inferredCategory !== exercise.category) {
      exercise.name = "";
      const nameInput = row.querySelector('[data-field="name"]');
      if (nameInput instanceof HTMLInputElement) {
        nameInput.value = "";
      }
    }
  }

  renderCategoryPicker(row, exercise.category);
  closeCategoryDrawer();

  if (openSuggestions) {
    openExerciseDrawer(row, exercise);
  }
}

function findExerciseMatches(category, query) {
  const source = category
    ? ALL_EXERCISE_DETAILS.filter((exercise) => exercise.category === category)
    : ALL_EXERCISE_DETAILS;
  const normalizedQuery = normalizeLookupValue(query);
  const profile = state.profile;

  if (!normalizedQuery) {
    return source
      .slice()
      .sort((left, right) => compareRecommendedExercises(left, right, profile));
  }

  return source
    .map((exercise) => ({
      exercise,
      score: scoreExerciseMatch(exercise, normalizedQuery),
      recommendation: getExerciseRecommendationScore(exercise, profile),
    }))
    .filter(({ score }) => score < Number.POSITIVE_INFINITY)
    .sort((left, right) => left.score - right.score || left.recommendation - right.recommendation || left.exercise.name.localeCompare(right.exercise.name))
    .map(({ exercise }) => exercise);
}

function compareRecommendedExercises(left, right, profile) {
  return getExerciseRecommendationScore(left, profile) - getExerciseRecommendationScore(right, profile)
    || left.name.localeCompare(right.name);
}

function getExerciseRecommendationScore(exercise, profile) {
  let score = 0;
  const levelOrder = { Iniciante: 0, Intermediario: 1, Avancado: 2 };
  const targetLevel = levelOrder[profile.level || ""];
  const exerciseLevel = levelOrder[exercise.level || ""];

  if (Number.isInteger(targetLevel) && Number.isInteger(exerciseLevel)) {
    score += Math.abs(targetLevel - exerciseLevel) * 2;
  }

  const goalAffinity = {
    Hipertrofia: ["Hipertrofia"],
    Emagrecimento: ["Emagrecimento", "Resistencia", "Condicionamento"],
    Resistencia: ["Resistencia", "Condicionamento"],
    Forca: ["Forca"],
    Reabilitacao: ["Reabilitacao", "Controle motor"],
  };
  const affinity = goalAffinity[profile.goal || ""] || [];

  if (affinity.length) {
    const hasAffinity = exercise.goals.some((item) => affinity.includes(item)) || affinity.includes(exercise.type);
    score += hasAffinity ? -3 : 2;
  }

  return score;
}

function scoreExerciseMatch(exercise, normalizedQuery) {
  const normalizedName = normalizeLookupValue(exercise.name);

  if (normalizedName.startsWith(normalizedQuery)) {
    return 0;
  }

  if (normalizedName.includes(normalizedQuery)) {
    return 1;
  }

  const aliasMatch = exercise.aliases.some((alias) => normalizeLookupValue(alias).includes(normalizedQuery));
  if (aliasMatch) {
    return 2;
  }

  const metadataFields = [
    exercise.level,
    exercise.type,
    exercise.equipment,
    exercise.description,
    exercise.postureTips,
    ...exercise.secondary,
    ...exercise.variations,
    ...exercise.goals,
  ];

  if (metadataFields.some((field) => normalizeLookupValue(field).includes(normalizedQuery))) {
    return 3;
  }

  return Number.POSITIVE_INFINITY;
}

function openExerciseDrawer(row, exercise) {
  if (!(row instanceof HTMLTableRowElement) || !exercise || !elements.exerciseDrawer) {
    return;
  }

  exerciseDrawerState.exerciseId = exercise.id;
  exerciseDrawerState.query = "";
  exerciseDrawerState.previewName = exercise.name || "";
  elements.exerciseDrawer.hidden = false;
  elements.exerciseDrawer.setAttribute("aria-hidden", "false");
  document.body.classList.add("drawer-open");
  renderExerciseDrawer();

  window.setTimeout(() => {
    elements.exerciseDrawerSearch?.focus();
    elements.exerciseDrawerSearch?.select();
  }, 0);
}

function openCategoryDrawer(row, exercise) {
  if (!(row instanceof HTMLTableRowElement) || !exercise || !elements.categoryDrawer) {
    return;
  }

  categoryDrawerState.exerciseId = exercise.id;
  elements.categoryDrawer.hidden = false;
  elements.categoryDrawer.setAttribute("aria-hidden", "false");
  document.body.classList.add("drawer-open");
  renderCategoryDrawer();
}

function closeExerciseDrawer() {
  if (!elements.exerciseDrawer || elements.exerciseDrawer.hidden) {
    return;
  }

  const currentExerciseId = exerciseDrawerState.exerciseId;
  elements.exerciseDrawer.hidden = true;
  elements.exerciseDrawer.setAttribute("aria-hidden", "true");
  document.body.classList.remove("drawer-open");
  exerciseDrawerState.exerciseId = null;
  exerciseDrawerState.query = "";
  exerciseDrawerState.previewName = "";

  const row = currentExerciseId
    ? elements.tableBody.querySelector(`tr[data-exercise-id="${currentExerciseId}"]`)
    : null;
  const nameInput = row?.querySelector('[data-field="name"]');

  if (nameInput instanceof HTMLInputElement) {
    nameInput.blur();
  }
}

function closeCategoryDrawer() {
  if (!elements.categoryDrawer || elements.categoryDrawer.hidden) {
    return;
  }

  elements.categoryDrawer.hidden = true;
  elements.categoryDrawer.setAttribute("aria-hidden", "true");
  categoryDrawerState.exerciseId = null;

  if (elements.exerciseDrawer.hidden) {
    document.body.classList.remove("drawer-open");
  }
}

function handleExerciseDrawerSearch(event) {
  exerciseDrawerState.query = event.target.value;
  renderExerciseDrawer();
}

function handleExerciseDrawerPreview(event) {
  const option = event.target instanceof Element ? event.target.closest('[data-action="choose-exercise-drawer"]') : null;

  if (!(option instanceof HTMLButtonElement)) {
    return;
  }

  exerciseDrawerState.previewName = option.dataset.exercise || "";
}

function renderExerciseDrawer() {
  if (!elements.exerciseDrawer || elements.exerciseDrawer.hidden) {
    return;
  }

  const exercise = getExerciseById(exerciseDrawerState.exerciseId);
  if (!exercise) {
    closeExerciseDrawer();
    return;
  }

  const activeTraining = getActiveTraining();
  const categoryLabel = exercise.category || "Todos os agrupamentos";
  const matches = findExerciseMatches(exercise.category, exerciseDrawerState.query);
  const visibleMatches = matches.slice(0, 40);
  const normalizedSelectedName = normalizeLookupValue(exercise.name);
  const selectedDetails = getExerciseDetails(exercise.name);

  if (elements.exerciseDrawerTraining) {
    elements.exerciseDrawerTraining.textContent = getDisplayTrainingName(activeTraining.name);
  }

  if (elements.exerciseDrawerCategory) {
    elements.exerciseDrawerCategory.textContent = categoryLabel;
  }

  if (elements.exerciseDrawerCategoryBadge) {
    elements.exerciseDrawerCategoryBadge.textContent = categoryLabel;
  }

  if (elements.exerciseDrawerDescription) {
    elements.exerciseDrawerDescription.textContent = exercise.category
      ? `Exibindo exercicios de ${exercise.category}. Busque tambem por nivel, equipamento, objetivo ou variacao.`
      : "Sem categoria definida. Busque por nome, equipamento, nivel ou escolha qualquer exercicio para preencher a linha.";
  }

  if (elements.exerciseDrawerSearch && elements.exerciseDrawerSearch.value !== exerciseDrawerState.query) {
    elements.exerciseDrawerSearch.value = exerciseDrawerState.query;
  }

  if (elements.exerciseDrawerMeta) {
    elements.exerciseDrawerMeta.textContent = `${categoryLabel} • ${matches.length} opcao(oes)`;
  }

  if (elements.exerciseDrawerMetaCopy) {
    const goal = state.profile.goal || "Livre";
    const level = state.profile.level || "Nao definido";
    elements.exerciseDrawerMetaCopy.textContent = `Ordenado por objetivo ${goal} e nivel ${level}.`;
  }

  if (!elements.exerciseDrawerList) {
    return;
  }

  if (!visibleMatches.length) {
    elements.exerciseDrawerList.innerHTML = '<div class="exercise-option exercise-option--empty">Nenhum exercicio encontrado.</div>';
    return;
  }

  elements.exerciseDrawerList.innerHTML = visibleMatches
    .map((exerciseEntry, index) => {
      const inferredCategory = exerciseEntry.category || inferCategoryFromExerciseName(exerciseEntry.name);
      const isSelected = normalizeLookupValue(exerciseEntry.name) === normalizedSelectedName
        || Boolean(selectedDetails && selectedDetails.name === exerciseEntry.name);
      return `<button class="exercise-option${isSelected ? " is-selected" : ""}" type="button" data-action="choose-exercise-drawer" data-exercise="${escapeAttribute(exerciseEntry.name)}" data-category="${escapeAttribute(inferredCategory)}"><span class="exercise-option__index">${String(index + 1).padStart(2, "0")}</span><span class="exercise-option__content"><strong>${exerciseEntry.name}</strong></span></button>`;
    })
    .join("");
}

function getRecommendationBadge(exercise, profile) {
  if (!exercise) {
    return "";
  }

  const matchesGoal = Boolean(profile.goal) && getExerciseRecommendationScore(exercise, profile) <= -1;
  const matchesLevel = Boolean(profile.level) && exercise.level === profile.level;

  if (matchesGoal && matchesLevel) {
    return "Perfil ideal";
  }

  if (matchesGoal) {
    return "Sugestao do objetivo";
  }

  if (matchesLevel) {
    return "Nivel alinhado";
  }

  return "";
}

function renderCategoryDrawer() {
  if (!elements.categoryDrawer || elements.categoryDrawer.hidden) {
    return;
  }

  const exercise = getExerciseById(categoryDrawerState.exerciseId);
  if (!exercise) {
    closeCategoryDrawer();
    return;
  }

  const activeTraining = getActiveTraining();
  const currentCategory = exercise.category || "Nao definido";

  if (elements.categoryDrawerTraining) {
    elements.categoryDrawerTraining.textContent = getDisplayTrainingName(activeTraining.name);
  }

  if (elements.categoryDrawerCurrent) {
    elements.categoryDrawerCurrent.textContent = currentCategory;
  }

  if (elements.categoryDrawerDescription) {
    elements.categoryDrawerDescription.textContent = "Escolha o agrupamento e, em seguida, selecione o exercicio no painel lateral.";
  }

  if (!elements.categoryDrawerList) {
    return;
  }

  elements.categoryDrawerList.innerHTML = Object.keys(EXERCISE_LIBRARY)
    .map((category) => {
      const isActive = category === exercise.category;
      const exerciseCount = EXERCISE_LIBRARY[category]?.length || 0;
      const shortCode = category.slice(0, 2).toUpperCase();
      return `<button class="category-option category-option--drawer${isActive ? " is-active" : ""}" type="button" data-action="choose-category-drawer" data-category="${category}"><span class="category-option__mark">${shortCode}</span><span class="category-option__body"><strong>${category}</strong></span><span class="category-option__meta">${exerciseCount} exercicio${exerciseCount === 1 ? "" : "s"}</span></button>`;
    })
    .join("");
}

function selectExerciseFromDrawer(optionButton) {
  const exercise = getExerciseById(exerciseDrawerState.exerciseId);
  const row = exerciseDrawerState.exerciseId
    ? elements.tableBody.querySelector(`tr[data-exercise-id="${exerciseDrawerState.exerciseId}"]`)
    : null;

  if (!exercise || !(row instanceof HTMLTableRowElement)) {
    closeExerciseDrawer();
    return;
  }

  exercise.name = optionButton.dataset.exercise || "";
  const nextCategory = optionButton.dataset.category || inferCategoryFromExerciseName(exercise.name);

  if (nextCategory && nextCategory !== exercise.category) {
    exercise.category = nextCategory;
    renderCategoryPicker(row, exercise.category);
  }

  const nameInput = row.querySelector('[data-field="name"]');
  if (nameInput instanceof HTMLInputElement) {
    nameInput.value = exercise.name;
    validateInputField(nameInput, "name", exercise.name);
  }

  persistAndRender(false);
  renderStats();
  closeExerciseDrawer();
  setStatusMessage(`Exercicio atualizado para ${exercise.name}.`, "success");
}

function selectCategoryFromDrawer(optionButton) {
  const exercise = getExerciseById(categoryDrawerState.exerciseId);
  const row = categoryDrawerState.exerciseId
    ? elements.tableBody.querySelector(`tr[data-exercise-id="${categoryDrawerState.exerciseId}"]`)
    : null;

  if (!exercise || !(row instanceof HTMLTableRowElement)) {
    closeCategoryDrawer();
    return;
  }

  setExerciseCategory(row, exercise, optionButton.dataset.category || "", false);
  persistAndRender(false);
  renderStats();
  closeCategoryDrawer();
  openExerciseDrawer(row, exercise);
  setStatusMessage(`Agrupamento atualizado para ${exercise.category}.`, "success");
}

function syncExerciseDrawerState() {
  if (!exerciseDrawerState.exerciseId) {
    return;
  }

  if (!getExerciseById(exerciseDrawerState.exerciseId)) {
    closeExerciseDrawer();
  }

  if (categoryDrawerState.exerciseId && !getExerciseById(categoryDrawerState.exerciseId)) {
    closeCategoryDrawer();
  }
}

function applySplitPreset(splitLabel, groups) {
  const activeTraining = getActiveTraining();
  const preset = SPLIT_PRESET_LIBRARY[splitLabel];
  const presetGroups = preset?.groups || groups;
  const insertedCount = preset ? replaceExercisesFromPreset(activeTraining, preset.exercises) : 0;

  activeTraining.splitLabel = splitLabel;
  activeTraining.focusGroups = presetGroups;
  persistAndRender();
  setStatusMessage(`Preset ${splitLabel} aplicado e treino atualizado com ${insertedCount} exercicio(s).`, "success");
}

function replaceExercisesFromPreset(training, presetExercises) {
  training.exercises = presetExercises.map((exercise) => ({
    id: crypto.randomUUID(),
    category: exercise.category,
    name: exercise.name,
    sets: exercise.sets,
    reps: exercise.reps,
    load: exercise.load,
    notes: "",
  }));

  return training.exercises.length;
}

function duplicateExercise(exerciseId) {
  const activeTraining = getActiveTraining();
  const exerciseIndex = activeTraining.exercises.findIndex((exercise) => exercise.id === exerciseId);

  if (exerciseIndex === -1) {
    return;
  }

  const source = activeTraining.exercises[exerciseIndex];
  const duplicate = {
    ...source,
    id: crypto.randomUUID(),
  };

  activeTraining.exercises.splice(exerciseIndex + 1, 0, duplicate);
  persistAndRender();
  setStatusMessage(`Exercicio ${source.name || "sem nome"} duplicado.`, "success");
}

function reorderExercise(sourceId, targetId) {
  const activeTraining = getActiveTraining();
  const sourceIndex = activeTraining.exercises.findIndex((exercise) => exercise.id === sourceId);
  const targetIndex = activeTraining.exercises.findIndex((exercise) => exercise.id === targetId);

  if (sourceIndex === -1 || targetIndex === -1) {
    return;
  }

  const [movedExercise] = activeTraining.exercises.splice(sourceIndex, 1);
  activeTraining.exercises.splice(targetIndex, 0, movedExercise);
  persistAndRender();
  setStatusMessage("Ordem dos exercicios atualizada.", "success");
}

function createCsv(payload) {
  const header = ["Treino", "Categoria", "Exercicio", "Series", "Repeticoes", "Carga", "Observacoes"];
  const rows = payload.flatMap((training) => {
    if (!training.exercicios.length) {
      return [[training.treino, "", "", "", "", "", ""]];
    }

    return training.exercicios.map((exercise) => [
      training.treino,
      exercise.category,
      exercise.name,
      exercise.sets,
      exercise.reps,
      exercise.load,
      exercise.notes,
    ]);
  });

  return [header, ...rows]
    .map((row) => row.map((value) => `"${String(value ?? "").replaceAll("\"", '""')}"`).join(";"))
    .join("\n");
}

function createExportPayload(includeAllTrainings = true) {
  const trainingsToExport = includeAllTrainings ? state.trainings : [getActiveTraining()];

  return trainingsToExport.map((training) => ({
    treino: getDisplayTrainingName(training.name),
    splitLabel: training.splitLabel || "Livre",
    profile: {
      ...createEmptyProfile(),
      ...(state.profile || {}),
    },
    exercicios: training.exercises.filter(isExerciseMeaningful).map((exercise) => ({
      category: exercise.category,
      name: exercise.name,
      sets: exercise.sets,
      reps: exercise.reps,
      load: exercise.load,
      notes: exercise.notes,
    })),
  }));
}

async function handleDocxDownload(payload, includeAllTrainings = true) {
  if (typeof docx === "undefined") {
    setStatusMessage("A biblioteca de exportacao DOCX nao foi carregada no navegador.", "error");
    return;
  }

  try {
    await gerarDocx(payload, includeAllTrainings ? "treino_personalizado_completo.docx" : "treino_personalizado.docx");
    setStatusMessage(
      includeAllTrainings
        ? "Todos os treinos exportados em DOCX profissional."
        : "Treino ativo exportado em DOCX profissional.",
      "success",
    );
  } catch (error) {
    setStatusMessage(error.message || "Falha ao gerar o arquivo DOCX.", "error");
  }
}

async function gerarDocx(dados = createExportPayload(false), fileName = "treino_personalizado.docx") {
  if (typeof docx === "undefined") {
    throw new Error("A biblioteca docx nao foi carregada no navegador.");
  }

  const blob = await gerarTreinoDocx(dados);
  downloadDocxBlob(blob, fileName);
  return blob;
}

function gerarPdf(dados = createExportPayload(false), fileName = "treino.pdf") {
  if (!window.jspdf?.jsPDF) {
    throw new Error("A biblioteca jsPDF nao foi carregada no navegador.");
  }

  const doc = new window.jspdf.jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
  });
  const profile = {
    ...createEmptyProfile(),
    ...(state.profile || {}),
  };
  const trainings = Array.isArray(dados) && dados.length ? dados : [{ treino: "Treino A", exercicios: [] }];
  const pageWidth = doc.internal.pageSize.getWidth();

  doc.setFillColor(18, 102, 86);
  doc.rect(12, 12, pageWidth - 24, 16, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("Treino de Academia", pageWidth / 2, 22, { align: "center" });

  doc.setTextColor(33, 33, 33);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");

  doc.autoTable({
    startY: 33,
    theme: "grid",
    margin: { left: 14, right: 14 },
    tableWidth: pageWidth - 28,
    styles: {
      font: "helvetica",
      fontSize: 9,
      cellPadding: 2,
      textColor: [33, 33, 33],
      lineColor: [180, 187, 184],
      lineWidth: 0.2,
    },
    columnStyles: {
      0: { fontStyle: "bold", fillColor: [245, 248, 247], cellWidth: 26 },
      1: { cellWidth: 66 },
      2: { fontStyle: "bold", fillColor: [245, 248, 247], cellWidth: 26 },
      3: { cellWidth: 66 },
    },
    body: [
      ["Aluno", profile.traineeName || "Nao informado", "Professor", profile.coachName || "Nao informado"],
      ["Objetivo", profile.goal || "Nao informado", "Nivel", profile.level || "Nao informado"],
      ["Data", profile.sessionDate ? formatDateForDocument(profile.sessionDate) : "Nao informada", "Duracao", profile.duration || "Nao informada"],
      ["Descanso", profile.rest || "Nao informado", "Frequencia", profile.frequency ? `${profile.frequency}x semana` : "Nao informada"],
    ],
  });

  let currentY = doc.lastAutoTable.finalY + 6;

  if (profile.notes) {
    const notesLines = doc.splitTextToSize(profile.notes, pageWidth - 36);
    doc.setFillColor(247, 248, 244);
    doc.rect(14, currentY, pageWidth - 28, 8 + (notesLines.length * 4.8), "F");
    doc.setDrawColor(180, 187, 184);
    doc.rect(14, currentY, pageWidth - 28, 8 + (notesLines.length * 4.8));
    doc.setFont("helvetica", "bold");
    doc.text("Observacoes gerais", 16, currentY + 5.5);
    doc.setFont("helvetica", "normal");
    doc.text(notesLines, 16, currentY + 11);
    currentY += 14 + (notesLines.length * 4.8);
  }

  trainings.forEach((training, index) => {
    if (currentY > 245) {
      doc.addPage();
      currentY = 18;
    } else if (index > 0) {
      currentY += 8;
    }

    doc.setFillColor(index % 2 === 0 ? 198 : 244, index % 2 === 0 ? 224 : 219, index % 2 === 0 ? 180 : 196);
    doc.rect(14, currentY, pageWidth - 28, 8, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text(`${training.treino || `Treino ${index + 1}`} • ${training.splitLabel || "Livre"}`, 16, currentY + 5.5);
    currentY += 10;

    const rows = (training.exercicios || []).map((exercise) => ([
      exercise.name || "Sem exercicio",
      exercise.sets || "-",
      exercise.reps || "-",
      exercise.load || "-",
      profile.rest || "-",
      extractTrainingDay(training.treino || ""),
    ]));

    doc.autoTable({
      startY: currentY,
      head: [["EXERCICIO", "SERIES", "REPETICOES", "CARGA", "PAUSA", "DIA"]],
      body: rows.length ? rows : [["Sem exercicios cadastrados", "-", "-", "-", "-", "-"]],
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: 9,
        cellPadding: 2.2,
        halign: "center",
        lineColor: [55, 55, 55],
        lineWidth: 0.2,
      },
      headStyles: {
        fillColor: [23, 70, 62],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      columnStyles: {
        0: { halign: "left" },
      },
      margin: { left: 14, right: 14 },
    });

    currentY = doc.lastAutoTable.finalY;
  });

  doc.save(ensurePdfFileName(fileName));
  return doc;
}

function handleXlsxDownload(payload, includeAllTrainings = true) {
  if (typeof XLSX === "undefined") {
    setStatusMessage("A biblioteca de exportacao XLSX nao foi carregada no navegador.", "error");
    return;
  }

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: "Planner de Treinos",
    Subject: "Fichas de treino",
    Author: "Planner de Treinos",
    CreatedDate: new Date(),
  };

  XLSX.utils.book_append_sheet(workbook, createSummaryWorksheet(payload), "Resumo");

  payload.forEach((training) => {
    const worksheet = createProfessionalWorksheet(training);
    XLSX.utils.book_append_sheet(workbook, worksheet, createWorksheetName(training.treino));
  });

  XLSX.writeFile(workbook, includeAllTrainings ? "treinos-completos.xlsx" : `${slugifyFileName(payload[0]?.treino || "treino")}.xlsx`);
  setStatusMessage(includeAllTrainings ? "Todos os treinos exportados em XLSX profissional." : "Treino ativo exportado em XLSX profissional.", "success");
}

function createSummaryWorksheet(payload) {
  const rows = [
    ["Planner de Treinos"],
    [`Gerado em ${new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(new Date())}`],
    ["Treino", "Divisao", "Total de exercicios", "Status"],
    ...payload.map((training) => [
      training.treino,
      training.splitLabel || "Livre",
      training.exercicios.length,
      training.exercicios.length ? "Preenchido" : "Vazio",
    ]),
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
  ];
  worksheet["!cols"] = [{ wch: 24 }, { wch: 18 }, { wch: 18 }, { wch: 16 }];
  worksheet["!rows"] = [{ hpt: 26 }, { hpt: 20 }, { hpt: 22 }];

  applyCellStyle(worksheet, "A1", {
    font: { name: "Arial", bold: true, sz: 18, color: { rgb: "17312A" } },
    fill: { fgColor: { rgb: "E4F2EC" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: createBorderStyle(),
  });
  applyCellStyle(worksheet, "A2", {
    font: { name: "Arial", italic: true, sz: 10, color: { rgb: "61736E" } },
    fill: { fgColor: { rgb: "F6F8F7" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: createBorderStyle(),
  });

  ["A3", "B3", "C3", "D3"].forEach((address) => {
    applyCellStyle(worksheet, address, {
      font: { name: "Arial", bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "245E53" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: createBorderStyle("FFFFFF"),
    });
  });

  for (let rowIndex = 3; rowIndex < rows.length; rowIndex += 1) {
    [0, 1, 2, 3].forEach((columnIndex) => {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      applyCellStyle(worksheet, address, {
        font: { name: "Arial", sz: 11, color: { rgb: "1F1F1F" } },
        fill: { fgColor: { rgb: rowIndex % 2 === 1 ? "FFFFFF" : "F5FAF8" } },
        alignment: { horizontal: columnIndex >= 2 ? "center" : "left", vertical: "center" },
        border: createBorderStyle(),
      });
    });
  }

  return worksheet;
}

function createProfessionalWorksheet(training) {
  const columns = ["Exercicio", "Series", "Repeticoes", "Carga (kg)", "Observacoes"];
  const profile = {
    ...createEmptyProfile(),
    ...(training.profile || {}),
  };
  const rows = [
    [training.treino],
    [`Divisao do dia: ${training.splitLabel || "Livre"}`],
    [`Aluno: ${profile.traineeName || "-"} | Objetivo: ${profile.goal || "-"} | Nivel: ${profile.level || "-"}`],
    [`Professor: ${profile.coachName || "-"} | Sessao: ${profile.sessionDate || "-"} | Duracao: ${profile.duration || "-"} | Descanso: ${profile.rest || "-"}`],
    columns,
  ];

  if (profile.notes) {
    rows.splice(4, 0, [`Observacoes profissionais: ${profile.notes}`]);
  }

  if (!training.exercicios.length) {
    rows.push(["Sem exercicios cadastrados", "", "", "", ""]);
  } else {
    training.exercicios.forEach((exercise) => {
      rows.push([exercise.name, exercise.sets, exercise.reps, exercise.load, exercise.notes]);
    });
  }

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const headerRowIndex = rows.findIndex((row) => row[0] === "Exercicio");
  worksheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 4 } },
  ];

  if (profile.notes) {
    worksheet["!merges"].push({ s: { r: 4, c: 0 }, e: { r: 4, c: 4 } });
  }

  worksheet["!cols"] = computeColumnWidths(rows);
  worksheet["!rows"] = [
    { hpt: 28 },
    { hpt: 20 },
    { hpt: 20 },
    { hpt: 20 },
    ...rows.slice(4).map(() => ({ hpt: 22 })),
  ];
  worksheet["!autofilter"] = { ref: `A${headerRowIndex + 1}:E${Math.max(rows.length, headerRowIndex + 2)}` };

  applyCellStyle(worksheet, "A1", {
    font: { name: "Arial", bold: true, sz: 16, color: { rgb: "17312A" } },
    fill: { fgColor: { rgb: "DFF0E8" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: createBorderStyle(),
  });
  ["A2", "A3", "A4"].forEach((address) => {
    applyCellStyle(worksheet, address, {
      font: { name: "Arial", italic: true, sz: 10, color: { rgb: "6A7974" } },
      fill: { fgColor: { rgb: address === "A4" ? "EEF5F2" : "F5F7F6" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: createBorderStyle(),
    });
  });

  if (profile.notes) {
    applyCellStyle(worksheet, "A5", {
      font: { name: "Arial", italic: true, sz: 10, color: { rgb: "506661" } },
      fill: { fgColor: { rgb: "F9FCFB" } },
      alignment: { horizontal: "left", vertical: "center", wrapText: true },
      border: createBorderStyle(),
    });
  }

  columns.forEach((_, columnIndex) => {
    const address = XLSX.utils.encode_cell({ r: headerRowIndex, c: columnIndex });
    applyCellStyle(worksheet, address, {
      font: { name: "Arial", bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "1C7A69" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: createBorderStyle("FFFFFF"),
    });
  });

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      applyCellStyle(worksheet, address, {
        font: { name: "Arial", sz: 11, color: { rgb: "1F1F1F" } },
        fill: { fgColor: { rgb: rowIndex % 2 === 1 ? "FFFFFF" : "F7FBFA" } },
        alignment: {
          horizontal: columnIndex > 0 && columnIndex < 4 ? "center" : "left",
          vertical: "center",
          wrapText: columnIndex === 4,
        },
        border: createBorderStyle(),
      });
    }
  }

  return worksheet;
}

async function gerarTreinoDocx(dados) {
  const {
    AlignmentType,
    BorderStyle,
    Document,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableLayoutType,
    TableRow,
    TextRun,
    VerticalAlign,
    WidthType,
  } = docx;

  const profile = {
    ...createEmptyProfile(),
    ...(state.profile || {}),
  };
  const generationDate = new Date();
  const generationLabel = new Intl.DateTimeFormat("pt-BR", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(generationDate);
  const sessionLabel = profile.sessionDate ? formatDateForDocument(profile.sessionDate) : formatDateForDocument(generationDate);
  const trainings = Array.isArray(dados) && dados.length ? dados : [{ treino: "Treino A", splitLabel: "Livre", exercicios: [] }];
  const sectionChildren = [
    new Paragraph({
      text: "Planilha de Treinamento Personalizado",
      spacing: { after: 120 },
    }),
    createDocxHeaderTable({
      AlignmentType,
      BorderStyle,
      Paragraph,
      Table,
      TableCell,
      TableLayoutType,
      TableRow,
      TextRun,
      VerticalAlign,
      WidthType,
    }),
    createDocxInfoTable({
      AlignmentType,
      BorderStyle,
      Paragraph,
      Table,
      TableCell,
      TableLayoutType,
      TableRow,
      TextRun,
      VerticalAlign,
      WidthType,
    }, profile, sessionLabel, generationLabel),
    new Paragraph({ text: " ", spacing: { after: 120 } }),
  ];

  trainings.forEach((training, index) => {
    sectionChildren.push(
      createDocxTrainingBlock({
        AlignmentType,
        BorderStyle,
        Paragraph,
        Table,
        TableCell,
        TableLayoutType,
        TableRow,
        TextRun,
        VerticalAlign,
        WidthType,
      }, training, index, profile),
    );

    if (index < trainings.length - 1) {
      sectionChildren.push(new Paragraph({ text: " ", spacing: { after: 140 } }));
    }
  });

  const document = new Document({
    creator: "Planner de Treinos",
    title: "Planilha de Treinamento Personalizado",
    description: "Ficha de treino profissional gerada no navegador",
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 720,
              right: 540,
              bottom: 720,
              left: 540,
            },
          },
        },
        children: sectionChildren,
      },
    ],
  });

  const generatedBlob = await Packer.toBlob(document);
  const arrayBuffer = await generatedBlob.arrayBuffer();
  return new Blob([arrayBuffer], { type: DOCX_MIME_TYPE });
}

function createDocxHeaderTable(docxApi) {
  const { AlignmentType, BorderStyle, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun, VerticalAlign, WidthType } = docxApi;

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: createDocxBorderSet(BorderStyle, "1C1C1C", 10),
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 1800, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "F4F8F7" },
            margins: { top: 120, right: 120, bottom: 120, left: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "AO", bold: true, size: 34, color: "1B7A6C", font: "Arial" }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "PERSONAL TRAINER", bold: true, size: 16, color: "86A96B", font: "Arial" }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 7600, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 240, right: 120, bottom: 240, left: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "PLANILHA DE TREINAMENTO PERSONALIZADO",
                    bold: true,
                    size: 28,
                    color: "1E1E1E",
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function createDocxInfoTable(docxApi, profile, sessionLabel, generationLabel) {
  const { AlignmentType, BorderStyle, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun, VerticalAlign, WidthType } = docxApi;
  const infoRows = [
    [
      createKeyValueCellDocx({ Paragraph, TextRun }, "PERSONAL TRAINER", profile.coachName || "Nao informado"),
      createKeyValueCellDocx({ Paragraph, TextRun }, "CREF", "Planejamento digital"),
    ],
    [
      createKeyValueCellDocx({ Paragraph, TextRun }, "ALUNO(A)", profile.traineeName || "Nao informado"),
      createKeyValueCellDocx({ Paragraph, TextRun }, "NIVEL", profile.level || "Nao informado"),
    ],
    [
      createKeyValueCellDocx({ Paragraph, TextRun }, "PAUSA", profile.rest || "Nao informado"),
      createKeyValueCellDocx({ Paragraph, TextRun }, "INICIO DO TREINAMENTO", sessionLabel),
    ],
    [
      createKeyValueCellDocx({ Paragraph, TextRun }, "PERIODO", profile.duration || "Nao informado"),
      createKeyValueCellDocx({ Paragraph, TextRun }, "FREQUENCIA", profile.frequency ? `${profile.frequency}x semanais` : "Nao informado"),
    ],
    [
      createWideInfoCellDocx({ Paragraph, TextRun }, "OBJETIVO", profile.goal || "Nao informado", 2),
    ],
    [
      createWideInfoCellDocx({ Paragraph, TextRun }, "FORMA DE TREINAMENTO", getTrainingMethodLabel(profile), 2),
    ],
    [
      createWideInfoCellDocx({ Paragraph, TextRun }, "REFERENCIAL", `Gerado automaticamente em ${generationLabel}`, 2),
    ],
  ];

  if (profile.notes) {
    infoRows.push([
      createWideInfoCellDocx({ Paragraph, TextRun }, "OBSERVACOES", profile.notes, 2),
    ]);
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
    rows: infoRows.map((cells) => new TableRow({ children: cells.map((cell) => new TableCell({
      width: { size: 4700, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      columnSpan: cell.columnSpan,
      shading: cell.shading,
      margins: { top: 70, right: 90, bottom: 70, left: 90 },
      borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
      children: cell.children,
    })) })),
  });
}

function createDocxTrainingBlock(docxApi, training, index, profile) {
  const { AlignmentType, BorderStyle, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun, VerticalAlign, WidthType } = docxApi;
  const headerFill = index % 2 === 0 ? "C7DEB5" : "F6DE8E";
  const warmupFill = index % 2 === 0 ? "E8F1DA" : "FCEAB6";
  const columns = [
    { label: "MAQ", width: 900 },
    { label: "EXERCICIOS", width: 3150 },
    { label: "SERIES", width: 950 },
    { label: "REPETICOES", width: 1150 },
    { label: "CARGA", width: 950 },
    { label: "PAUSA", width: 900 },
    { label: "DIA", width: 850 },
  ];
  const dayLabel = extractTrainingDay(training.treino);
  const exerciseRows = training.exercicios.length
    ? training.exercicios
    : [{ category: "", name: "Sem exercicios cadastrados", sets: "-", reps: "-", load: "-", notes: "" }];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
    rows: [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 7,
            shading: { fill: headerFill },
            borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
            margins: { top: 70, right: 90, bottom: 70, left: 90 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `${training.treino.toUpperCase()}: ${buildTrainingSubtitle(training, profile)}`,
                    bold: true,
                    size: 20,
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 7,
            shading: { fill: warmupFill },
            borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
            margins: { top: 60, right: 90, bottom: 60, left: 90 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "AQUECIMENTO: MOBILIDADE", bold: true, size: 18, font: "Arial" }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: columns.map((column) => new TableCell({
          width: { size: column.width, type: WidthType.DXA },
          shading: { fill: headerFill },
          verticalAlign: VerticalAlign.CENTER,
          borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
          margins: { top: 40, right: 40, bottom: 40, left: 40 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: column.label, bold: true, size: 16, font: "Arial" })],
            }),
          ],
        })),
      }),
      ...exerciseRows.map((exercise) => new TableRow({
        children: [
          createTrainingValueCellDocx(docxApi, exercise.category || "-", columns[0].width, AlignmentType.CENTER),
          createExerciseNameCellDocx(docxApi, exercise, columns[1].width),
          createTrainingValueCellDocx(docxApi, exercise.sets || "-", columns[2].width, AlignmentType.CENTER),
          createTrainingValueCellDocx(docxApi, exercise.reps || "-", columns[3].width, AlignmentType.CENTER),
          createTrainingValueCellDocx(docxApi, exercise.load || "-", columns[4].width, AlignmentType.CENTER),
          createTrainingValueCellDocx(docxApi, profile.rest || "-", columns[5].width, AlignmentType.CENTER),
          createTrainingValueCellDocx(docxApi, dayLabel, columns[6].width, AlignmentType.CENTER),
        ],
      })),
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 7,
            borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
            margins: { top: 60, right: 90, bottom: 60, left: 90 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "RESFRIAMENTO: Alongamentos.", bold: true, size: 18, font: "Arial" }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function slugifyFileName(value) {
  return String(value || "treino")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "treino";
}

function createDocxBorderSet(borderStyle, color = "1C1C1C", size = 8) {
  return {
    top: { style: borderStyle.SINGLE, color, size },
    right: { style: borderStyle.SINGLE, color, size },
    bottom: { style: borderStyle.SINGLE, color, size },
    left: { style: borderStyle.SINGLE, color, size },
  };
}

function createKeyValueCellDocx(docxApi, label, value) {
  const { Paragraph, TextRun } = docxApi;

  return {
    columnSpan: 1,
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: `${label}: `, bold: true, size: 17, font: "Arial" }),
          new TextRun({ text: value, size: 17, font: "Arial" }),
        ],
      }),
    ],
  };
}

function createWideInfoCellDocx(docxApi, label, value, columnSpan = 2) {
  const { Paragraph, TextRun } = docxApi;

  return {
    columnSpan,
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: `${label}: `, bold: true, size: 17, font: "Arial" }),
          new TextRun({ text: value, size: 17, font: "Arial" }),
        ],
      }),
    ],
  };
}

function createTrainingValueCellDocx(docxApi, value, width, alignment) {
  const { BorderStyle, Paragraph, TableCell, TextRun, VerticalAlign, WidthType } = docxApi;

  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
    margins: { top: 45, right: 45, bottom: 45, left: 45 },
    children: [
      new Paragraph({
        alignment,
        children: [
          new TextRun({ text: String(value || "-"), size: 17, font: "Arial" }),
        ],
      }),
    ],
  });
}

function createExerciseNameCellDocx(docxApi, exercise, width) {
  const { AlignmentType, BorderStyle, Paragraph, TableCell, TextRun, VerticalAlign, WidthType } = docxApi;

  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: createDocxBorderSet(BorderStyle, "1C1C1C", 8),
    margins: { top: 45, right: 60, bottom: 45, left: 60 },
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({ text: exercise.name || "Sem exercicio", bold: true, size: 17, font: "Arial" }),
        ],
      }),
      ...(exercise.notes
        ? [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: `Obs.: ${exercise.notes}`, italics: true, size: 15, color: "5C5C5C", font: "Arial" }),
              ],
            }),
          ]
        : []),
    ],
  });
}

function formatDateForDocument(value) {
  const date = value instanceof Date ? value : new Date(`${value}T00:00:00`);

  if (Number.isNaN(date.getTime())) {
    return new Intl.DateTimeFormat("pt-BR", { dateStyle: "short" }).format(new Date());
  }

  return new Intl.DateTimeFormat("pt-BR", { dateStyle: "short" }).format(date);
}

function getTrainingMethodLabel(profile) {
  return [
    profile.goal && `Foco em ${profile.goal.toLowerCase()}`,
    profile.level && `nivel ${profile.level.toLowerCase()}`,
    profile.duration && `sessao de ${profile.duration}`,
  ].filter(Boolean).join(" • ") || "Prescricao personalizada de musculacao";
}

function buildTrainingSubtitle(training, profile) {
  return [
    training.splitLabel,
    training.exercicios[0]?.category,
    profile.goal,
  ].filter(Boolean).join(" | ") || "Treino personalizado";
}

function extractTrainingDay(trainingName) {
  const match = String(trainingName || "").match(/[A-Z]$/i);
  return match ? match[0].toUpperCase() : "-";
}

function applyCellStyle(worksheet, address, style) {
  if (!worksheet[address]) {
    worksheet[address] = { t: "s", v: "" };
  }

  worksheet[address].s = style;
}

function createBorderStyle(color = "CAD9D5") {
  return {
    top: { style: "thin", color: { rgb: color } },
    right: { style: "thin", color: { rgb: color } },
    bottom: { style: "thin", color: { rgb: color } },
    left: { style: "thin", color: { rgb: color } },
  };
}

function computeColumnWidths(rows) {
  const widths = [];

  rows.forEach((row) => {
    row.forEach((value, columnIndex) => {
      const contentLength = String(value ?? "").length;
      widths[columnIndex] = Math.max(widths[columnIndex] || 12, contentLength + 4);
    });
  });

  return widths.map((width, columnIndex) => ({
    wch: columnIndex === 4 ? Math.min(Math.max(width, 26), 44) : Math.min(Math.max(width, 14), 28),
  }));
}

function downloadFile(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = fileName;
  anchor.click();

  URL.revokeObjectURL(url);
}

function downloadDocxBlob(blob, fileName) {
  const safeFileName = ensureDocxFileName(fileName);
  const docxBlob = blob instanceof Blob && blob.type === DOCX_MIME_TYPE
    ? blob
    : new Blob([blob], { type: DOCX_MIME_TYPE });
  const url = URL.createObjectURL(docxBlob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = safeFileName;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);
}

function ensureDocxFileName(fileName) {
  const normalizedName = String(fileName || "treino_personalizado.docx").trim() || "treino_personalizado.docx";
  return normalizedName.toLowerCase().endsWith(".docx") ? normalizedName : `${normalizedName}.docx`;
}

function ensurePdfFileName(fileName) {
  const normalizedName = String(fileName || "treino.pdf").trim() || "treino.pdf";
  return normalizedName.toLowerCase().endsWith(".pdf") ? normalizedName : `${normalizedName}.pdf`;
}

async function parseImportFile(file) {
  const extension = file.name.split(".").pop()?.toLowerCase();

  if (extension === "json") {
    const text = await file.text();
    return normalizeImportedTrainings(parseJsonContent(text));
  }

  if (extension === "csv") {
    const text = await file.text();
    return normalizeImportedTrainings(parseCsvContent(text));
  }

  throw new Error("Formato de importacao nao suportado. Use JSON ou CSV.");
}

function parseJsonContent(text) {
  const parsed = JSON.parse(text);

  if (Array.isArray(parsed)) {
    return parsed;
  }

  if (Array.isArray(parsed.trainings)) {
    return parsed.trainings;
  }

  throw new Error("JSON invalido para importacao de treinos.");
}

function parseCsvContent(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2) {
    throw new Error("CSV sem dados suficientes para importacao.");
  }

  const rows = lines.map(parseCsvLine);
  const trainingsMap = new Map();

  rows.slice(1).forEach((columns) => {
    const hasCategoryColumn = columns.length >= 7;
    const trainingName = columns[0];
    const category = hasCategoryColumn ? columns[1] : "";
    const exerciseName = hasCategoryColumn ? columns[2] : columns[1];
    const sets = hasCategoryColumn ? columns[3] : columns[2];
    const reps = hasCategoryColumn ? columns[4] : columns[3];
    const load = hasCategoryColumn ? columns[5] : columns[4];
    const notes = hasCategoryColumn ? columns[6] : columns[5];
    const normalizedTrainingName = (trainingName || "Treino importado").trim();

    if (!trainingsMap.has(normalizedTrainingName)) {
      trainingsMap.set(normalizedTrainingName, {
        treino: normalizedTrainingName,
        exercicios: [],
      });
    }

    if ([category, exerciseName, sets, reps, load, notes].every((value) => !String(value || "").trim())) {
      return;
    }

    trainingsMap.get(normalizedTrainingName).exercicios.push({
      category: category || inferCategoryFromExerciseName(exerciseName || "") || "",
      name: exerciseName || "",
      sets: sets || "",
      reps: reps || "",
      load: load || "",
      notes: notes || "",
    });
  });

  return Array.from(trainingsMap.values());
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let isInsideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"') {
      if (isInsideQuotes && nextChar === '"') {
        current += '"';
        index += 1;
      } else {
        isInsideQuotes = !isInsideQuotes;
      }
      continue;
    }

    if (char === ";" && !isInsideQuotes) {
      result.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  result.push(current);
  return result;
}

function normalizeImportedTrainings(rawTrainings) {
  const trainings = rawTrainings
    .map((training, index) => normalizeTraining(training, index))
    .filter(Boolean);

  if (!trainings.length) {
    throw new Error("Os dados importados nao possuem treinos utilizaveis.");
  }

  return trainings;
}

function normalizeTraining(training, index) {
  const rawName = training?.treino ?? training?.name ?? `Treino ${index + 1}`;
  const rawExercises = training?.exercicios ?? training?.exercises ?? [];

  return {
    id: crypto.randomUUID(),
    name: String(rawName).slice(0, 30),
    splitLabel: String(training?.splitLabel ?? ""),
    focusGroups: Array.isArray(training?.focusGroups) ? training.focusGroups.filter(Boolean) : [],
    exercises: Array.isArray(rawExercises)
      ? rawExercises.map((exercise) => normalizeExercise(exercise)).filter(Boolean)
      : [],
  };
}

function normalizeExercise(exercise) {
  if (!exercise || typeof exercise !== "object") {
    return null;
  }

  const normalizedName = String(exercise.name ?? exercise.exercicio ?? "");

  return {
    id: crypto.randomUUID(),
    category: String(exercise.category ?? exercise.categoria ?? inferCategoryFromExerciseName(normalizedName) ?? ""),
    name: normalizedName,
    sets: String(exercise.sets ?? exercise.series ?? ""),
    reps: String(exercise.reps ?? exercise.repeticoes ?? ""),
    load: String(exercise.load ?? exercise.carga ?? ""),
    notes: String(exercise.notes ?? exercise.observacoes ?? ""),
  };
}

function inferCategoryFromExerciseName(name) {
  const normalizedName = normalizeLookupValue(name);

  if (!normalizedName) {
    return "";
  }

  return EXERCISE_CATEGORY_BY_NAME.get(normalizedName)
    || ALL_EXERCISE_DETAILS.find((exercise) => normalizeLookupValue(exercise.name).includes(normalizedName))?.category
    || "";
}

function getExerciseDetails(name) {
  const normalizedName = normalizeLookupValue(name);

  if (!normalizedName) {
    return null;
  }

  return EXERCISE_DETAILS_BY_NAME.get(normalizedName)
    || ALL_EXERCISE_DETAILS.find((exercise) => normalizeLookupValue(exercise.name).includes(normalizedName))
    || null;
}

function normalizeLookupValue(value) {
  return String(value || "").trim().toLowerCase();
}

function escapeAttribute(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function isExerciseMeaningful(exercise) {
  return [exercise.category, exercise.name, exercise.sets, exercise.reps, exercise.load, exercise.notes]
    .some((value) => String(value || "").trim());
}

function getNextTrainingName() {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const nextIndex = state.trainings.length;
  return `Treino ${alphabet[nextIndex] || nextIndex + 1}`;
}

function createWorksheetName(name) {
  return name.replace(/[\\/*?:\[\]]/g, " ").slice(0, 31) || "Treino";
}

function getPrintSubtitle(training) {
  const profile = {
    ...createEmptyProfile(),
    ...(state.profile || {}),
  };

  return [
    profile.traineeName && `Aluno: ${profile.traineeName}`,
    profile.goal && `Objetivo: ${profile.goal}`,
    profile.level && `Nivel: ${profile.level}`,
    training.splitLabel && `Divisao: ${training.splitLabel}`,
  ].filter(Boolean).join(" • ");
}

function updatePrintMetadata() {
  const activeTraining = getActiveTraining();
  const printTitle = getDisplayTrainingName(activeTraining.name);
  const printSubtitle = getPrintSubtitle(activeTraining);

  document.title = `${printTitle} - Planner de Treinos`;
  document.body.dataset.printTitle = printTitle;
  document.body.dataset.printSubtitle = printSubtitle;

  if (elements.briefingBoard) {
    elements.briefingBoard.dataset.printTitle = printTitle;
    elements.briefingBoard.dataset.printSubtitle = printSubtitle;
  }
}

function formatTimestamp(timestamp) {
  const date = new Date(timestamp);

  return new Intl.DateTimeFormat("pt-BR", {
    hour: "2-digit",
    minute: "2-digit",
    day: "2-digit",
    month: "2-digit",
  }).format(date);
}

function getDisplayTrainingName(name) {
  return name.trim() || "Treino sem nome";
}

function setStatusMessage(message, tone = "neutral") {
  elements.statusMessage.textContent = message;
  elements.statusMessage.classList.toggle("is-error", tone === "error");
  elements.statusMessage.classList.toggle("is-success", tone === "success");
}

function persistAndRender(shouldRenderAll = true) {
  saveState();
  updatePrintMetadata();

  if (shouldRenderAll) {
    render();
  }
}
