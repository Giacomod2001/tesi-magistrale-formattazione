Introduzione

Ogni anno le imprese europee investono miliardi di euro in pubblicità digitale per presidiare mercati sempre più frammentati e competitivi. Secondo i dati dell'Interactive Advertising Bureau (IAB Europe, 2024), la spesa in digital advertising nell'area UE ha superato i 96 miliardi di euro nel 2023, confermando un tasso di crescita annuo composto prossimo all'11 % nell'ultimo quinquennio. Parallelamente, il settore delle agenzie per il lavoro destina al canale digitale una quota crescente del proprio budget promozionale, nella convinzione — spesso non verificata — che la visibilità online si traduca linearmente in candidature qualificate e, in ultima analisi, in ricavi.

Eppure, nonostante l'entità degli investimenti, la domanda fondamentale rimane largamente inevasa: quale euro produce valore, e quale viene dissipato?

La questione non è nuova. Già John Wanamaker — con la celebre sentenza attribuita anche a Lord Leverhulme — lamentava che «metà del denaro speso in pubblicità è sprecato; il problema è che non so quale metà». A distanza di oltre un secolo, l'ipertrofia dei canali digitali ha moltiplicato, anziché ridurre, la complessità del problema: i decision-maker operano oggi in un ecosistema caratterizzato da molteplicità di touchpoint, latenza variabile degli effetti pubblicitari, interazioni non lineari tra canali e asimmetrie informative strutturali tra piattaforme e inserzionisti.

Il presente lavoro si colloca esattamente in questo snodo critico. L'obiettivo è duplice: da un lato, progettare e implementare una pipeline di Marketing Mix Modeling (MMM) calibrata sulle specificità del settore delle agenzie per il lavoro — un dominio finora trascurato dalla letteratura accademica — utilizzando il caso Randstad Italia come banco di prova; dall'altro, esplorare un paradigma decisionale che la tesi definisce human-in-the-middle, nel quale l'output algoritmico non sostituisce la discrezionalità del marketing manager, ma la informa, la struttura e, auspicabilmente, la migliora.

La scelta del MMM come strumento analitico non è casuale. Nell'era della progressiva erosione dei meccanismi di tracciamento deterministico — dal declino dei cookie di terze parti all'introduzione dell'App Tracking Transparency di Apple — il Marketing Mix Modeling sta conoscendo una seconda giovinezza, candidandosi a divenire il perno di un nuovo paradigma di misurazione privacy-first (Chan & Perry, 2017; Jin et al., 2017). A differenza dei modelli di attribuzione multi-touch, che dipendono dal tracciamento a livello di utente, il MMM opera su dati aggregati — spesa pubblicitaria, impression, variabili contestuali — e non necessita di identificatori individuali, risultando così strutturalmente compatibile con i vincoli regolatori contemporanei.

Il settore delle agenzie per il lavoro, e in particolare il caso Randstad Italia, offre un terreno di indagine particolarmente stimolante. La variabile obiettivo è multilivello — dall'impressione pubblicitaria alla candidatura, dal colloquio all'inserimento lavorativo — e il ciclo di conversione è lungo, variabile e soggetto a una ciclicità macroeconomica pronunciata. I canali di recruiting digitale — Indeed, LinkedIn, Google, Meta — operano con logiche di pricing, targeting e misurazione radicalmente diverse, rendendo la comparazione e l'ottimizzazione cross-canale una sfida tanto analitica quanto manageriale.

La tesi propone infine un modello decisionale — il paradigma human-in-the-middle — nel quale il decisore non è relegato né al ruolo di validatore passivo dell'output algoritmico, né a quello di utilizzatore discrezionale che può ignorare le raccomandazioni a piacimento, ma opera come un nodo attivo all'interno di un sistema di feedback nel quale la sua conoscenza contestuale e il suo giudizio qualitativo informano, e sono a loro volta informati da, le evidenze quantitative del modello.


Struttura del lavoro

L'indagine si articola in due parti principali.

La Parte I — Il Problema (Capitoli 1–2) delimita il perimetro della ricerca. Il Capitolo 1 ricostruisce lo stato dell'arte del Marketing Mix Modeling, esaminandone l'evoluzione metodologica — dal declino dell'attribuzione deterministica all'ascesa dei framework open-source — e identifica le lacune bibliografiche che motivano l'indagine: l'assenza di applicazioni documentate nel settore HR e la limitata analisi dell'integrazione tra output algoritmico e discrezionalità del decisore. Il Capitolo 2 contestualizza il caso Randstad Italia, esaminando le dinamiche del mercato del lavoro italiano, i canali strategici del recruiting digitale e le criticità operative affrontate quotidianamente dai manager.

La Parte II — La Soluzione (Capitolo 3) formula la proposta metodologica. Il Capitolo 3 espone i fondamenti matematici dei modelli di Marketing Mix Modeling, chiarendo le differenze tra l'approccio frequentista e quello bayesiano, e giustificando l'adozione della metodologia alla luce delle sfide delineate.


Nota sui dati. I dati utilizzati nella presente indagine sono sintetici, calibrati sulle distribuzioni reali di Randstad Italia. La scelta è motivata dai vincoli di compliance aziendale vigenti al momento della stesura. L'integrazione di dati reali aggregati sarà valutata in fase successiva, compatibilmente con le autorizzazioni necessarie.
