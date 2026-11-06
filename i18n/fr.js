
/*
 * SimpleSPC French translations
 *
 * Initial French dictionary.
 * Terminology will be reviewed with a French-speaking
 * SPC practitioner before the bilingual release.
 */

window.SimpleSPCi18n.registerLanguage("fr", {

  // Translation and locale polish
  "app.browserTitle": "Outil SPC simple",
  "data.chooseFile": "Choisir un fichier",
  "data.noFileChosen": "Aucun fichier sélectionné",
  "common.formatLetter.bold": "G",
  "common.formatLetter.italic": "I",
  "common.formatLetter.underline": "S",
  "chart.columns.tDatesTooltip": "Non utilisée pour une carte T fondée sur les dates des événements.",
  "errors.annotationRequired": "Veuillez saisir une date et un libellé pour l’annotation.",
  "errors.noData": "Aucune donnée chargée. Importez un CSV ou utilisez d’abord l’éditeur de données.",
  "errors.tMinEvents": "La carte T nécessite au moins 4 événements.",
  "errors.gMinPoints": "La carte G nécessite au moins 4 points (chaque valeur doit être supérieure ou égale à 1).",
  "errors.gPositive": "Les valeurs de la carte G doivent être des nombres supérieurs ou égaux à 1.",
  "errors.gProbability": "Impossible de calculer la probabilité de la carte G à partir de ces données (vérifiez que les valeurs sont ≥ 1).",
  "errors.axisBounds": "Le minimum de l’axe Y ne peut pas être supérieur au maximum.",
  "errors.tDatesAxis": "La carte T utilisant les dates des événements nécessite un axe Date / heure.",
  "errors.tooFewValidPoints": "Pas assez de points valides après lecture des données. Vérifiez les colonnes sélectionnées.",
  "errors.tDateModeAxis": "La carte T nécessite un axe Date / heure (elle utilise les dates des événements).",
  "errors.tGapsMin": "La carte T (mode intervalles) nécessite au moins 3 valeurs d’intervalle valides.",
  "errors.chartNotImplemented": "Le type de graphique « {chartType} » n’est pas encore disponible.",
  "errors.cMinPoints": "La carte C nécessite au moins 2 points.",
  "errors.pMinPoints": "La carte P nécessite au moins 2 points.",
  "errors.pPrimeMinPoints": "La carte P′ nécessite au moins 2 points.",
  "errors.uMinPoints": "La carte U nécessite au moins 2 points.",
  "errors.clipboardUnsupported": "La copie dans le presse-papiers n’est pas prise en charge par ce navigateur. Utilisez plutôt « Enregistrer le(s) graphique(s) sous… ».",
  "errors.exportImageFailed": "Désolé, votre navigateur n’a pas pu exporter cette image.",
  "errors.imageCreationFailed": "Impossible de créer l’image.",
  "errors.projectColumnsMissing": "Les paramètres du projet ont été appliqués, mais certaines colonnes enregistrées sont introuvables. Veuillez resélectionner : {columns}",
  "errors.portableUnrecognised": "Ce fichier ne contient pas de projet portable FreeSPC reconnu.",
  "errors.portableRead": "Impossible de lire le graphique SPC enregistré. Le fichier peut être endommagé ou incompatible. Détail technique : {message}",
  "errors.portableLoaderMissing": "Impossible de charger le graphique SPC enregistré car le chargeur de projet n’est pas disponible.",
  "errors.portableLoad": "Impossible de charger le graphique SPC enregistré. Détail technique : {message}",
  "errors.portableStart": "Impossible de démarrer correctement le fichier FreeSPC enregistré. Détail technique : {message}",

  // Application heading
  "app.title": "Outil Web Simple SPC",

  "app.subtitle": "Importez des données et créez des cartes de contrôle statistique directement dans votre navigateur. Vos données ne quittent jamais votre navigateur.",

  // Main buttons
  "buttons.generate": "Générer / Recalculer",
  "buttons.reset": "Réinitialiser le graphique",
  "buttons.help": "Aide et conseils",
  "buttons.spcHelper": "Assistant SPC",

  // Main tabs
  "tabs.data": "Données",
  "tabs.chart": "Graphique",
  "tabs.labels": "Libellés",
  "tabs.axes": "Axes",
  "tabs.baseline": "Référence",
  
  "tabs.rules": "Règles",

  // Sélecteur de langue
  "language.label": "Langue",

  // Introduction
  "app.intro": "Ajoutez vos données, choisissez votre graphique et ses paramètres, puis cliquez sur « Générer / Recalculer » si nécessaire. Vous pouvez importer des fichiers CSV ou Excel, modifier les données dans l'éditeur et enregistrer votre graphique pour le rouvrir plus tard. Vous avez des questions sur la maîtrise statistique des procédés ? Demandez à l'assistant.",

  // Infobulles
  "buttons.resetTitle": "Conserver les données chargées et les colonnes sélectionnées, mais réinitialiser les paramètres du graphique, notamment les objectifs, les changements de phase, les annotations, les règles et la mise en forme.",

  
"tabs.baselineTitle": "Définir une période de référence initiale ou un objectif de performance",

// PANNEAU DONNÉES
"data.addHeading": "Ajouter des données :",
"data.requirementsAria": "Aide sur les données nécessaires",
"data.requirementsTooltip": "Les colonnes nécessaires dépendent du type de graphique. La plupart des cartes SPC nécessitent une colonne de temps ou de séquence et une mesure. Certaines nécessitent également des nombres d'événements et des dénominateurs. Les diagrammes de Pareto nécessitent une colonne de catégories et peuvent également utiliser une colonne d'effectifs.",
"data.uploadHint": "Importez un fichier CSV/Excel ou collez vos données dans l'éditeur.",
"data.openEditor": "Ouvrir l'éditeur de données (coller/saisir)",
"data.clear": "Effacer les données et le graphique",
"data.clearTitle": "Supprimer les données chargées et réinitialiser le graphique. Enregistrez votre travail au préalable si vous souhaitez le rouvrir plus tard.",

"data.saveHeading": "Enregistrer ou ouvrir un graphique",
"data.saveHint": "Enregistrez votre graphique, vos données et vos paramètres dans un fichier JSON pour pouvoir les rouvrir et les modifier ultérieurement.",
"data.saveButton": "Enregistrer le graphique",
"data.openButton": "Ouvrir un graphique enregistré",

"data.portableHeading": "Enregistrer une copie hors ligne",
"data.portableAria": "Aide sur les graphiques portables",
"data.portableTooltip": "Enregistrez une copie HTML autonome de votre graphique. Ce fichier contient vos données et une version fonctionnelle de SimpleSPC. Vous pouvez l'ouvrir sans connexion Internet, apporter des modifications et l'enregistrer à nouveau. Les fichiers portables contiennent vos données. Stockez-les et partagez-les conformément aux exigences de gouvernance de l'information de votre organisation.",
"data.portableButton": "Enregistrer un graphique portable",

// PANNEAU GRAPHIQUE
"chart.noData": "Chargez des données pour activer le choix du graphique et des colonnes.",
"chart.newDataHeading": "Nouvelles données chargées",
"chart.newDataQuestion": "Souhaitez-vous être guidé pour choisir le graphique le plus adapté à ces données ?",
"chart.helpChoose": "Aidez-moi à choisir",
"chart.chooseManually": "Choisir manuellement",
"chart.typeHeading": "Type de graphique :",
"chart.run": "Graphique chronologique",
"chart.xmr": "Carte XmR",
"chart.moreTypes": "Autres types de graphiques",

"chart.setup": "Paramètres du graphique",

// CONTRÔLES DES CARTES SECONDAIRES MR / S
"chart.secondary.showMR": "Afficher la carte MR",
"chart.secondary.showS": "Afficher la carte S",
"chart.secondary.mrDisplay": "Affichage MR",
"chart.secondary.sDisplay": "Affichage de la carte S",
"chart.secondary.includesSplits": "(phases incluses)",
"chart.secondary.current": "Période actuelle",
"chart.secondary.all": "Toutes",
"chart.secondary.helpAria": "Aide sur l'affichage de la carte secondaire",
"chart.secondary.helpText": "Les cartes XmR peuvent afficher une carte des étendues mobiles (MR), qui permet d'étudier la variation entre deux mesures consécutives. Les cartes X̄–S peuvent afficher une carte S, qui représente la variation au sein des sous-groupes. Période actuelle affiche uniquement la dernière période d'analyse. Toutes affiche l'historique complet, y compris les différentes phases. Utilisez la case à cocher pour afficher ou masquer la carte secondaire.",

// SÉLECTION DES COLONNES
"chart.columns.heading": "Choisir les colonnes :",
"chart.columns.xDefault": "Date / colonne de l'axe X",
"chart.columns.yDefault": "Valeur / colonne de l'axe Y",
"chart.columns.thirdDefault": "Dénominateur / opportunités",
"chart.columns.xmrY": "Mesure (utilisée pour les limites XmR)",
"chart.columns.cY": "Nombre (c) par période",
"chart.columns.pY": "Numérateur : unités non conformes (d)",
"chart.columns.pThird": "Dénominateur : total (n)",
"chart.columns.pHint": "La carte P représente une proportion : d sur n.",
"chart.columns.pprimeHint": "La carte P′ représente une proportion : d sur n, avec des limites de contrôle ajustées selon la méthode de Laney.",
"chart.columns.uY": "Numérateur : non-conformités (c)",
"chart.columns.uThird": "Dénominateur : opportunités (n)",
"chart.columns.uHint": "La carte U représente le nombre de non-conformités par opportunité : c sur n.",
"chart.columns.uprimeHint": "La carte U′ représente le nombre de non-conformités par opportunité, avec des limites de contrôle ajustées selon la méthode de Laney.",
"chart.columns.xbarsY": "Valeur mesurée",
"chart.columns.xbarsThird": "Identifiant du sous-groupe (jour / semaine / échantillon)",
"chart.columns.xbarsHint": "La carte X̄–S nécessite plusieurs mesures par sous-groupe. L'axe horizontal utilise automatiquement l'ordre des sous-groupes.",
"chart.columns.tGapsY": "Temps entre les événements (par ex. en jours)",
"chart.columns.tDatesY": "Colonne de valeurs non utilisée (la carte T utilise les dates)",
"chart.columns.tGapsHint": "La carte T représente le temps entre des événements rares à partir des intervalles numériques fournis.",
"chart.columns.tDatesHint": "La carte T représente le temps entre des événements rares, calculé à partir des dates des événements.",
"chart.columns.gY": "Opportunités entre les événements",
"chart.columns.gHint": "La carte G représente le nombre d'opportunités entre des événements rares.",
"chart.columns.paretoX": "Colonne des catégories",

"chart.columns.paretoY": "Colonne des fréquences / effectifs",

// AUTRES TYPES DE GRAPHIQUES
"chart.types.c": "Carte C (dénombrements)",
"chart.types.p": "Carte P (proportions)",
"chart.types.pprime": "Carte P′ (Laney)",
"chart.types.u": "Carte U (taux)",
"chart.types.uprime": "Carte U′ (Laney)",
"chart.types.xbars": "Carte X̄–S (sous-groupes)",
"chart.types.t": "Carte T (temps entre événements)",
"chart.types.g": "Carte G (opportunités entre événements)",
"chart.types.pareto": "Diagramme de Pareto (catégories)",

// AFFICHAGE DES TAUX
"chart.rate.label": "Afficher le taux pour :",
"chart.rate.per1": "Pour 1",
"chart.rate.per10": "Pour 10",
"chart.rate.per100": "Pour 100",
"chart.rate.per1000": "Pour 1 000",
"chart.rate.per10000": "Pour 10 000",
"chart.rate.hint": "Modifie uniquement l'affichage du taux ; le calcul sous-jacent de la carte U reste inchangé.",

// TYPE D'AXE HORIZONTAL
"chart.axis.type": "Type d'axe horizontal",
"chart.axis.helpAria": "Aide sur le type d'axe horizontal",
"chart.axis.help": "Utilisez Date / heure pour les cartes SPC chronologiques. Utilisez Séquence / catégorie pour afficher les observations dans leur ordre initial sans interpréter les valeurs de l'axe horizontal comme des dates.",
"chart.axis.date": "Date / heure",
"chart.axis.sequence": "Séquence / catégorie",

// FORMAT DES DATES
"chart.dateFormat.label": "Format des dates numériques",
"chart.dateFormat.helpAria": "Aide sur le format des dates",
"chart.dateFormat.help": "Utilisé pour les dates numériques ambiguës, comme 01/02/2024. Les dates au format ISO, comme 2024-02-01, et les dates Excel sont reconnues automatiquement.",
"chart.dateFormat.uk": "Royaume-Uni / jour d'abord (jj/mm/aaaa)",
"chart.dateFormat.us": "États-Unis / mois d'abord (mm/jj/aaaa)",
"chart.dateFormat.auto": "Détection automatique",

"chart.dateFormat.iso": "Dates ISO / Excel uniquement",

// DISPONIBILITÉ DES CARTES
"chart.availability.noData": "Chargez des données pour déterminer les types de cartes disponibles.",
"chart.availability.selectedColumn": "la colonne de valeurs sélectionnée",
"chart.availability.xmr": "Moins de 12 points de données sont disponibles dans « {column} ». Un minimum de 12 points est nécessaire pour une carte XmR valide.",
"chart.availability.c": "Aucune colonne appropriée de nombres entiers non négatifs n'est disponible. Les cartes C nécessitent des dénombrements par période.",
"chart.availability.p": "Aucune paire numérateur/dénominateur appropriée n'est disponible. Les cartes P nécessitent deux colonnes de nombres entiers, le numérateur représentant une partie du dénominateur.",
"chart.availability.pprime": "Aucune paire numérateur/dénominateur appropriée n'est disponible. Les cartes P′ nécessitent deux colonnes de nombres entiers, le numérateur représentant une partie du dénominateur.",
"chart.availability.u": "Aucune paire nombre d'événements/opportunités appropriée n'est disponible. Les cartes U nécessitent une colonne de dénombrements entiers et une colonne d'opportunités strictement positives.",
"chart.availability.uprime": "Aucune paire nombre d'événements/opportunités appropriée n'est disponible. Les cartes U′ nécessitent une colonne de dénombrements entiers et une colonne d'opportunités strictement positives.",
"chart.availability.xbars": "Les cartes X̄–S nécessitent une colonne de mesures numériques et une colonne de sous-groupes contenant des identifiants répétés.",
"chart.availability.t": "Les cartes T nécessitent des dates d'événements ou une colonne numérique indiquant le temps entre les événements.",
"chart.availability.g": "Les cartes G nécessitent une colonne de nombres entiers dont toutes les valeurs sont au moins égales à 1, car elles représentent le nombre d'opportunités entre des événements rares.",

// AVERTISSEMENTS SUR LES DATES
"chart.warning.mixedDates": "Attention : cette colonne semble contenir un mélange de dates numériques au format britannique et américain. Harmonisez les formats si possible.",
"chart.warning.usDatesUkSelected": "Attention : ces dates semblent utiliser le format américain, mois/jour/année (mm/jj/aaaa), alors que le format britannique, jour/mois/année, est sélectionné. Certaines lignes pourraient être ignorées. Essayez le format américain ou la détection automatique.",
"chart.warning.ukDatesUsSelected": "Attention : ces dates semblent utiliser le format britannique, jour/mois/année (jj/mm/aaaa), alors que le format américain, mois/jour/année, est sélectionné. Certaines lignes pourraient être ignorées. Essayez le format britannique ou la détection automatique.",
"chart.warning.ambiguousUk": "Cette colonne contient des dates numériques ambiguës. Elles seront actuellement interprétées au format britannique, jour/mois/année (jj/mm/aaaa).",
"chart.warning.ambiguousUs": "Cette colonne contient des dates numériques ambiguës. Elles seront actuellement interprétées au format américain, mois/jour/année (mm/jj/aaaa).",
"chart.warning.ambiguousAuto": "Cette colonne contient des dates numériques ambiguës et l'application ne peut pas déterminer leur format avec suffisamment de certitude. Elle utilisera le format britannique par défaut, sauf si vous choisissez une autre option.",
"chart.warning.ambiguousIso": "Cette colonne contient des dates numériques avec des barres obliques ou des traits d'union. En mode « Dates ISO / Excel uniquement », les dates numériques ambiguës peuvent ne pas être reconnues comme des dates.",

// CONTRÔLE DU FORMAT DES DATES DÉSACTIVÉ

"chart.dateFormat.disabled": "Le choix du format des dates est disponible uniquement lorsque l'axe horizontal utilise Date / heure.",

// LIBELLÉS COMMUNS DE MISE EN FORME
"common.autoDefault": "Auto / par défaut",
"common.auto": "Auto",
"common.bold": "Gras",
"common.italic": "Italique",
"common.underline": "Souligné",

// PANNEAU LIBELLÉS
"labels.heading": "Libellés du graphique",
"labels.title": "Titre du graphique",
"labels.xAxis": "Libellé de l'axe X",
"labels.yAxis": "Libellé de l'axe Y",
"labels.titleSettings": "Paramètres du titre",
"labels.fontFamily": "Police du titre",
"labels.fontSize": "Taille de police du titre",
"labels.showTitle": "Afficher le titre du graphique",

// PANNEAU AXES
"axes.heading": "Axes",
"axes.xSection": "Axe X",
"axes.ySection": "Axe Y",
"axes.xFontFamily": "Police de l'axe X",
"axes.xFontSize": "Taille de police de l'axe X",
"axes.yMin": "Minimum de l'axe Y",
"axes.yMax": "Maximum de l'axe Y",
"axes.yRangeAria": "Aide sur l'échelle de l'axe Y",
"axes.yRangeHelp": "Laissez les valeurs minimale et maximale vides pour que le graphique détermine automatiquement son échelle.",
"axes.yFontFamily": "Police de l'axe Y",
"axes.yFontSize": "Taille de police de l'axe Y",
"axes.yFormat": "Format de l'axe Y",
"axes.formatAuto": "Auto",
"axes.formatDecimal": "Décimal",
"axes.formatInteger": "Entier",
"axes.formatPercent": "Pourcentage",
"axes.decimals": "Nombre de décimales de l'axe Y",
"axes.decimalsAuto": "Auto",
"axes.dp0": "0 décimale",
"axes.dp1": "1 décimale",
"axes.dp2": "2 décimales",
"axes.dp3": "3 décimales",
"axes.tickInterval": "Intervalle des graduations de l'axe Y",
"axes.tickAria": "Aide sur l'intervalle des graduations",
"axes.tickHelp": "Laissez ce champ vide pour un espacement automatique. Pour les pourcentages, indiquez des points de pourcentage : par exemple, 1 correspond à 1 %, et 0.5 correspond à 0,5 %.",

// PANNEAU PÉRIODE DE RÉFÉRENCE
"baseline.heading": "Données de référence",
"baseline.helpAria": "Aide sur la période de référence",
"baseline.help": "Une période de référence est une période initiale utilisée pour comprendre le fonctionnement du processus avant les observations suivantes. Indiquez le nombre de points à utiliser pour établir cette référence. SimpleSPC utilise ces points pour calculer la ligne centrale initiale et, le cas échéant, les limites de contrôle. Les observations suivantes restent visibles pour permettre la comparaison. Laissez ce champ vide si vous n'avez pas besoin d'une période de référence initiale distincte. Choisissez cette période en fonction de votre connaissance du processus, et non simplement pour éliminer des signaux gênants.",
"baseline.useFirst": "Utiliser les",
"baseline.pointsAsBaseline": "premiers points comme période de référence",
"baseline.blankHint": "(laisser vide si aucune période de référence)",
"baseline.targetHeading": "Objectif",
"baseline.targetHelpAria": "Aide pour définir un objectif",
"baseline.targetHelp": "Indiquez le niveau de performance que vous souhaitez atteindre, puis précisez si les valeurs supérieures ou égales, ou inférieures ou égales, à cet objectif sont souhaitables. Un objectif est un niveau de performance visé, et non une limite de contrôle statistique.",
"baseline.targetHint": "Indiquez votre valeur cible, puis choisissez si des valeurs plus élevées ou plus faibles sont préférables.",
"baseline.targetValue": "Valeur cible",
"baseline.atOrAbove": "Une valeur supérieure ou égale à l'objectif est favorable",
"baseline.atOrBelow": "Une valeur inférieure ou égale à l'objectif est favorable",
"baseline.hideTarget": "Masquer la ligne d'objectif",

"baseline.showTarget": "Afficher la ligne d'objectif",

// PANNEAU DES RÈGLES
"rules.heading": "Règles de détection",
"rules.explainer": "Explication des règles",

"rules.shift": "Longueur d'une série (décalage)",
"rules.shiftAria": "Aide sur la règle de décalage",
"rules.shiftHelp": "Nombre de points consécutifs du même côté de la ligne centrale nécessaires pour détecter un décalage. Valeur par défaut : 8.",

"rules.trend": "Longueur de la tendance",
"rules.trendAria": "Aide sur la règle de tendance",
"rules.trendHelp": "Nombre d'augmentations ou de diminutions consécutives nécessaires pour détecter une tendance. Valeur par défaut : 6. Ce paramètre est utilisé uniquement lorsque la règle de tendance est disponible et activée.",

"rules.advanced": "Options de détection avancées",
"rules.enableTrend": "Activer la règle de tendance",
"rules.enableTrendHelp": "Recherche une succession prolongée d'augmentations ou de diminutions consécutives. La longueur requise est définie par le paramètre Longueur de la tendance ci-dessus.",
"rules.advancedCaution": "Option avancée — peut augmenter le nombre de fausses alertes.",

"rules.rareLabel": "Règles de séries et de tendances pour les cartes T/G",
"rules.rareAria": "Aide sur les règles pour événements rares",
"rules.rareHelp": "Active la détection des séries et des tendances sur les cartes T et G pour événements rares. Ces types de cartes présentent naturellement des irrégularités ; les règles de détection de motifs doivent donc être utilisées avec prudence.",
"rules.rareCaution": "Option avancée — les cartes d'événements rares peuvent produire de faux signaux de séries ou de tendances.",

"rules.clamp": "Limiter la LIC à zéro",
"rules.clampAria": "Aide sur la limite inférieure de contrôle",
"rules.clampHelp": "Si la limite inférieure de contrôle calculée est négative, afficher la LIC à zéro à la place. Utile pour les mesures telles que les dénombrements ou les durées, qui ne peuvent pas être négatives.",

"rules.zoneIntro": "Règles de zones facultatives (désactivées par défaut). À utiliser uniquement pour les types de cartes qui les prennent en charge.",
"rules.zone23": "Appliquer la règle « 2 points sur 3 dans le tiers extérieur »",
"rules.zone45": "Appliquer la règle « 4 points sur 5 au-delà de 1 sigma »",
"rules.zoneCaution": "Les règles supplémentaires de détection de motifs peuvent augmenter les fausses alertes sur les cartes XmR / individuelles. Elles sont désactivées par défaut et ne doivent être utilisées qu'en comprenant ce compromis.",

"rules.mrHeading": "Calcul des limites XmR",
"rules.mrAria": "Aide sur le filtrage des étendues mobiles XmR",
"rules.mrHelp": "Une variation très importante entre deux points consécutifs peut rendre les limites de contrôle XmR trop larges. Lorsque cette option est activée, SimpleSPC conserve tous les points et toutes les étendues mobiles sur le graphique, mais n'utilise pas les étendues mobiles anormalement grandes pour estimer la variation habituelle. La méthode de Nelson est utilisée : les étendues mobiles supérieures à 3,5 fois la moyenne initiale des étendues mobiles sont exclues de l'estimation de la variation.",
"rules.mrScreen": "Exclure les variations exceptionnellement grandes du calcul des limites",

"rules.conservative": "Cette carte utilise un ensemble de règles prudent, conçu pour permettre une interprétation fiable.",
"rules.flag": "Signaler les causes spéciales sur le graphique",
"rules.colourBlind": "Couleurs et symboles adaptés aux personnes daltoniennes",

// INFOBULLES DYNAMIQUES
"rules.disabledTrend": "La règle de tendance n'est pas proposée pour ce type de carte.",
"rules.enableTrendFirst": "Cochez « Activer la règle de tendance » pour utiliser ce paramètre.",
"rules.enableRareFirst": "Activez les règles de séries et de tendances pour ce type de carte afin d'utiliser ce paramètre.",
"rules.rareOnly": "Utilisé uniquement pour les cartes T et G.",
"rules.zoneBlocked": "Les règles de zones ne sont pas disponibles pour ce type de carte, car elles pourraient produire des alertes trompeuses.",

// CONFIRMATIONS POUR LES ÉVÉNEMENTS RARES
"rules.rareConfirmImmediate": "Les cartes T et G présentent naturellement des irrégularités.\n\nLes règles de séries et de tendances peuvent produire de faux signaux sur ce type de carte et sont désactivées par défaut.\n\nN'activez ces règles que si vous comprenez ce compromis.",

"rules.rareConfirmOnce": "Option avancée : règles de séries et de tendances sur les cartes {chart}\n\nCes motifs apparaissent souvent par hasard sur les cartes T/G et peuvent augmenter les fausses alertes.\n\nActiver malgré tout ?",

// FENÊTRE D'EXPLICATION DES RÈGLES
"explainer.title": "Explication des règles SPC",
"explainer.close": "Fermer",
"explainer.closeAria": "Fermer l'explication des règles",
"explainer.subtitle": "{chart} : cet outil utilise des règles prudentes, adaptées au type de carte, afin de limiter les fausses alertes.",

// NOMS DES CARTES
"explainer.chart.run": "Graphique chronologique",
"explainer.chart.xmr": "Carte XmR",
"explainer.chart.c": "Carte C",
"explainer.chart.p": "Carte P",
"explainer.chart.pprime": "Carte P′",
"explainer.chart.u": "Carte U",
"explainer.chart.uprime": "Carte U′",
"explainer.chart.xbars": "Carte X̄–S",
"explainer.chart.t": "Carte T",
"explainer.chart.g": "Carte G",
"explainer.chart.pareto": "Diagramme de Pareto",

// POURQUOI LES RÈGLES DIFFÈRENT
"explainer.why.attribute": "Les cartes pour données attributaires utilisent ici un ensemble de règles prudent. Les règles de zones sont désactivées, car elles peuvent produire des alertes trompeuses, notamment lorsque les limites varient ou que les dénombrements sont faibles.",
"explainer.why.rare": "Les cartes d'événements rares présentent naturellement des irrégularités. Les règles de séries et de tendances peuvent produire de faux signaux ; leur activation est donc réservée aux options avancées et accompagnée d'un avertissement.",
"explainer.why.xmr": "Les cartes XmR suivent des mesures individuelles dans le temps. SimpleSPC conserve volontairement une interprétation simple par défaut, tout en proposant des règles supplémentaires de détection de motifs et des options de calcul pour les utilisateurs qui en ont besoin.",
"explainer.why.xbars": "Les cartes X̄–S constituent la famille de cartes pour laquelle des règles avancées sont les plus justifiables, mais les règles activées par défaut restent prudentes pour limiter les fausses alertes.",
"explainer.why.run": "Les graphiques chronologiques n'utilisent pas de limites de contrôle. L'outil privilégie par défaut une interprétation simple et prudente.",
"explainer.why.pareto": "Un diagramme de Pareto n'est pas une carte de contrôle ordonnée dans le temps : les règles de détection des signaux SPC ne s'y appliquent donc pas. Il classe les catégories pour orienter les investigations et les actions d'amélioration.",
"explainer.why.general": "Cet outil privilégie par défaut un ensemble de règles simple et prudent afin de limiter les fausses alertes.",

// RÈGLES APPLICABLES
"explainer.checksHeading": "Ce que cette carte vérifie par défaut",
"explainer.rule.beyond": "Points au-delà des limites",
"explainer.rule.run": "Règle des séries",
"explainer.rule.trend": "Règle de tendance",
"explainer.rule.zone23": "2 points sur 3 dans le tiers extérieur",
"explainer.rule.zone45": "4 points sur 5 au-delà de 1 sigma",
"explainer.status.on": "appliquée par défaut.",
"explainer.status.optional": "disponible dans les options avancées, désactivée par défaut.",
"explainer.status.warn": "réservée aux options avancées, désactivée par défaut, avec un avertissement avant utilisation.",
"explainer.status.unavailable": "non proposée pour ce type de carte.",

// POURQUOI CERTAINES RÈGLES SONT LIMITÉES
"explainer.limitedHeading": "Pourquoi certaines règles sont limitées",
"explainer.limited1": "Cet outil est conçu pour favoriser une interprétation fiable dans les secteurs de la santé et des services publics. Il privilégie un nombre réduit de signaux plus fiables plutôt qu'un ensemble de règles produisant davantage de fausses alertes.",
"explainer.limited2": "Un « signal » peut entraîner des réunions, des inquiétudes, des investigations ou des actions. Pour réduire les efforts inutiles et la fatigue liée aux alertes, certaines règles sont réservées aux types de cartes pour lesquels elles sont les plus fiables.",

// DÉFINITIONS DES RÈGLES
"explainer.definitionsHeading": "Définitions simples des règles",
"explainer.definition.beyond": "<strong>Points au-delà des limites :</strong> un point au-dessus de la limite supérieure ou au-dessous de la limite inférieure de contrôle.",
"explainer.definition.run": "<strong>Règle des séries :</strong> une suite prolongée de points du même côté de la ligne centrale.",
"explainer.definition.trend": "<strong>Règle de tendance :</strong> une succession prolongée d'augmentations ou de diminutions consécutives.",
"explainer.definition.zone": "<strong>Règles de zones :</strong> règles supplémentaires fondées sur la distance des points par rapport à la ligne centrale. Leur applicabilité est plus limitée ; elles ne sont proposées que pour les types de cartes où leur usage est le mieux justifié.",

// EXPLICATION TECHNIQUE DES CARTES XMR
"explainer.xmr.heading": "Comment sont calculées les limites XmR",
"explainer.xmr.intro": "Les cartes XmR estiment la variation habituelle à partir des différences entre les points consécutifs, appelées <strong>étendues mobiles</strong>.",
"explainer.xmr.largeJump": "Un écart très important entre deux points peut surestimer la variation et produire des limites de contrôle trop larges. Par défaut, FreeSPC écarte les étendues mobiles exceptionnellement grandes de l'estimation utilisée pour calculer les limites XmR. Les points et les étendues mobiles d'origine restent visibles sur la carte.",
"explainer.xmr.technical": "<strong>Détail technique :</strong> FreeSPC utilise la méthode de Nelson. Les étendues mobiles supérieures à 3,5 fois la moyenne initiale des étendues mobiles ne sont pas utilisées pour estimer la variation habituelle.",
"explainer.xmr.current": "<strong>Paramètre actuel :</strong> le filtrage des étendues mobiles est <strong>{state}</strong>. Vous pouvez modifier ce réglage dans les <em>Options de détection avancées</em>.",
"explainer.xmr.split": "<strong>Que se passe-t-il lorsque vous ajoutez un changement de phase ?</strong> FreeSPC traite les données avant et après ce point comme deux périodes distinctes du processus. Chaque période possède sa propre moyenne et ses propres limites de contrôle ; l'étendue mobile qui franchit la séparation n'est pas utilisée.",
"explainer.xmr.minimum": "FreeSPC nécessite au moins <strong>{points} points</strong> pour interpréter les limites XmR d'une période. Si un changement de phase crée une période plus courte, les limites sont présentées comme provisoires et FreeSPC ne conclut pas encore que cette période présente uniquement une variation habituelle ou des causes spéciales.",
"explainer.on": "ACTIVÉ",
"explainer.off": "DÉSACTIVÉ",

// ZONE DES RÉSULTATS
"results.chartHeading": "Graphique :",
"results.chartHelpAria": "Aide sur le menu contextuel du graphique",
"results.chartHelp": "Faites un clic droit sur le graphique pour accéder à des options supplémentaires. Selon le type de carte, vous pouvez ajouter des annotations, créer ou supprimer des changements de phase, exclure ou réintégrer des points étudiés dans le calcul des limites, gérer les libellés des lignes centrales et copier ou exporter votre graphique. Certaines actions nécessitent un clic droit directement sur un point de données.",

// GUIDE DE DÉMARRAGE
"results.quickStart": "Démarrage rapide :",
"results.quick1": "Importez un fichier CSV ou Excel, ou ouvrez l'éditeur de données pour coller ou saisir des données",
"results.quick2Before": "Si nécessaire, corrigez les données ou supprimez les lignes et colonnes inutiles dans l'éditeur, puis cliquez sur",
"results.quickApply": "Appliquer les données",
"results.quick3": "Choisissez le type de carte et vérifiez les colonnes X et Y",
"results.quick4Before": "Cliquez sur",
"results.quick4After": "si les modifications ne sont pas prises en compte automatiquement",
"results.quick5": "Ajoutez des titres, des libellés, des objectifs, des changements de phase ou des annotations selon vos besoins",
"results.quick6Before": "Faites un clic droit sur le graphique pour accéder aux options supplémentaires et utilisez",
"results.quick6After": "si vous souhaitez reprendre votre travail ultérieurement",

// CONTRÔLES DYNAMIQUES
"results.hideControls": "Masquer les paramètres",
"results.showControls": "Afficher les paramètres",
"results.needsRecalc": "Modifications enregistrées — cliquez sur Générer / Recalculer",

// TITRES DES CARTES SECONDAIRES
"results.mrHeading": "Carte des étendues mobiles :",
"results.sHeading": "Carte S :",

"results.mrCanvasTitle": "Étendues mobiles (MR)",

// LÉGENDES DES CARTES
"visual.value": "Valeur",
"visual.centreLine": "Ligne centrale",
"visual.median": "Médiane",
"visual.ucl": "LSC",
"visual.lcl": "LIC",
"visual.target": "Objectif",
"visual.xbar": "X̄",
"visual.count": "Effectif",
"visual.cumulative": "Pourcentage cumulé",
"visual.movingRange": "Étendue mobile",
"visual.mrAverage": "Moyenne des étendues mobiles",
"visual.mrUcl": "LSC des étendues mobiles",
"visual.mrAveragePeriod": "Moyenne des étendues mobiles (période {period})",
"visual.mrUclPeriod": "LSC des étendues mobiles (période {period})",
"visual.omittedExcluded": "Non utilisée : adjacente à un point exclu",
"visual.omittedScreened": "Non utilisée : étendue mobile exceptionnellement grande",

// INFOBULLES DE LA CARTE DES ÉTENDUES MOBILES
"visual.mrTooltip": "Étendue mobile : {value}",
"visual.mrExcludedTooltip": "Non utilisée pour estimer la variation : {value} (adjacente à un point exclu)",
"visual.mrScreenedTooltip": "Non utilisée pour estimer la variation : {value} (étendue mobile exceptionnellement grande, écartée du calcul)",

// LIBELLÉS SUPPLÉMENTAIRES XMR
"visual.mean": "Moyenne",
"visual.ucl3sigma": "LSC (3σ)",
"visual.lcl3sigma": "LIC (3σ)",

// INFOBULLES DU DIAGRAMME DE PARETO
"visual.pareto.count": "Effectif : {count} ({share} % du total)",
"visual.pareto.cumulative": "Pourcentage cumulé : {percent} %",

"visual.pareto.through": "Pourcentage cumulé jusqu'à cette catégorie : {percent} %",

// TITRES AUTOMATIQUES DES CARTES
"auto.title.run": "Graphique chronologique",
"auto.title.xmr": "Carte X (valeurs individuelles)",
"auto.title.c": "Carte C",
"auto.title.p": "Carte P",
"auto.title.pprime": "Carte P′",
"auto.title.u": "Carte U",
"auto.title.uprime": "Carte U′",
"auto.title.t": "Carte T",
"auto.title.tGaps": "Carte T (intervalles)",
"auto.title.g": "Carte G",
"auto.title.pareto": "Diagramme de Pareto",
"auto.title.xbar": "Carte X̄",
"auto.title.s": "Carte S",

// LIBELLÉS AUTOMATIQUES DES AXES
"auto.axis.date": "Date",
"auto.axis.value": "Valeur",
"auto.axis.category": "Catégorie",
"auto.axis.count": "Effectif",
"auto.axis.subgroup": "Sous-groupe",
"auto.axis.proportion": "Proportion / %",
"auto.axis.rate": "Taux par opportunité",
"auto.axis.rateN": "Taux pour {n} opportunités",
"auto.axis.opportunities": "Opportunités entre événements",
"auto.axis.time": "Temps entre événements",
"auto.axis.timeDays": "Temps entre événements (jours)",
"auto.axis.cumulative": "Pourcentage cumulé",
"auto.axis.s": "S",

// INTRODUCTIONS DES RÉSUMÉS STATISTIQUES
"summary.heading": "Résumé",
"summary.xbarsHeading": "Résumé de la carte X̄–S",
"summary.chart.run": "Graphique chronologique",
"summary.chart.xmr": "Carte XmR",

"summary.totalPoints": "Nombre total de points :",
"summary.totalSubgroups": "Nombre total de sous-groupes :",
"summary.divided": "Le graphique est divisé en",
"summary.period": "période",
"summary.periods": "périodes",
"summary.basedBaseline": "(selon la période de référence et les éventuels changements de phase).",
"summary.basedSplits": "(selon les changements de phase définis).",

"summary.rateIntro": "Taux affichés pour",
"summary.opportunity": "opportunité",
"summary.opportunities": "opportunités",

"summary.note": "Remarque :",
"summary.rareNote": "Les cartes d'événements rares présentent souvent des distributions asymétriques ; les limites de contrôle peuvent donc ne pas être symétriques comme celles d'une carte XmR.",

"summary.paretoOccurrences": "Nombre total d'occurrences :",
"summary.paretoAcross": "réparties dans",
"summary.paretoCategories": "catégories.",

// TITRES DES PÉRIODES ET COUVERTURE
"summary.periodLabel": "Période",
"summary.singlePeriod": "Période unique",
"summary.initialSegment": "(segment initial / période de référence)",
"summary.of": "sur",
"summary.coverage": "Période couverte :",
"summary.point": "point",
"summary.points": "points",
"summary.subgroups": "sous-groupes",
"summary.to": "à",

// MESURES STATISTIQUES
"summary.mean": "Moyenne",
"summary.medianPeriod": "Médiane (cette période)",
"summary.centreLine": "Ligne centrale",
"summary.controlLimits": "limites de contrôle",
"summary.lowerLimit": "LIC",
"summary.upperLimit": "LSC",
"summary.range": "intervalle",
"summary.estimatedSigma": "σ estimé (à partir des étendues mobiles)",
"summary.averageMr": "moyenne des étendues mobiles",
"summary.centreGap": "Ligne centrale (intervalle moyen entre événements)",
"summary.centreOpportunities": "Ligne centrale (nombre moyen d'opportunités)",
"summary.upperLimitLabel": "limite supérieure",
"summary.centreAndLimits": "Ligne centrale et limites",
"summary.notAvailable": "(non disponible)",

// MESURES DES CARTES X-BAR / S
"summary.subgroupSize": "Taille du sous-groupe",
"summary.measurement": "mesure",
"summary.measurements": "mesures",
"summary.perSubgroup": "par sous-groupe",
"summary.variesSubgroup": "(variable selon le sous-groupe)",
"summary.notStated": "non précisé",
"summary.xbarCentre": "Ligne centrale de la carte X̄",
"summary.sCentre": "Ligne centrale de la carte S",

"summary.estimatedProcessSd": "Écart-type estimé du processus",

// EXPLICATIONS DE LA PÉRIODE DE RÉFÉRENCE X-BAR / S
"summary.xbars.baseline": "Période de référence",
"summary.xbars.firstOne": "le premier sous-groupe (1)",
"summary.xbars.firstMany": "les",
"summary.xbars.subgroupsMany": "premiers sous-groupes",
"summary.xbars.windowOne": "définit la période utilisée pour le calcul des limites ;",
"summary.xbars.windowMany": "définissent la période utilisée pour le calcul des limites ;",
"summary.xbars.usedEstimate": "sous-groupes pris en compte pour estimer les limites",
"summary.xbars.excludedOne": "exclu",
"summary.xbars.excludedMany": "exclus",

// EXPLICATIONS DES LIMITES DE CONTRÔLE X-BAR / S
"summary.xbars.limitsVary": "les limites de contrôle varient selon la taille du sous-groupe",
"summary.xbars.sLimitsVary": "La valeur attendue de S et ses limites de contrôle varient selon la taille du sous-groupe.",
"summary.xbars.centreRange": "Intervalle des lignes centrales",
"summary.xbars.lclRange": "Intervalle des LIC",
"summary.xbars.uclRange": "Intervalle des LSC",
"summary.xbars.limits": "limites",

// ÉTAT DES CARTES
"summary.xbars.xChart": "Carte X̄",
"summary.xbars.sChart": "Carte S",
"summary.xbars.interpretation": "Interprétation",
"summary.xbars.latestPeriod": "Dernière période",
"summary.xbars.stableX": "stable (aucun signal clair de changement dans les moyennes des sous-groupes).",
"summary.xbars.stableS": "stable (aucun signal clair de changement dans la variation au sein des sous-groupes).",
"summary.xbars.signals": "signal(s) :",
"summary.xbars.signalDetected": "un ou plusieurs signaux de cause spéciale ont été détectés.",

// INTERPRÉTATIONS COMBINÉES PAR PÉRIODE
"summary.xbars.bothStable": "Les moyennes des sous-groupes (X̄) et la variation au sein des sous-groupes (S) semblent toutes deux stables pendant cette période.",
"summary.xbars.xUnstable": "Les moyennes des sous-groupes (X̄) présentent des signaux de cause spéciale, tandis que la variation au sein des sous-groupes (S) semble stable.",
"summary.xbars.sUnstable": "Les moyennes des sous-groupes (X̄) semblent stables, tandis que la variation au sein des sous-groupes (S) présente des signaux de cause spéciale.",
"summary.xbars.bothUnstable": "Les moyennes des sous-groupes (X̄) et la variation au sein des sous-groupes (S) présentent toutes deux des signaux de cause spéciale.",

// ÉVALUATION DE LA DERNIÈRE PÉRIODE
"summary.xbars.latestBothStable": "La dernière période semble stable sur les deux parties de la carte X̄–S. Cela signifie que le niveau actuel du processus et la variation au sein des sous-groupes présentent un comportement globalement cohérent selon les règles de détection actuellement sélectionnées.",
"summary.xbars.latestXUnstable": "La dernière période présente un changement possible dans les moyennes des sous-groupes, tandis que la variation au sein des sous-groupes reste stable.",
"summary.xbars.latestSUnstable": "La dernière période présente un changement possible dans la variation au sein des sous-groupes, tandis que les moyennes des sous-groupes restent stables.",
"summary.xbars.latestBothUnstable": "La dernière période présente des signes possibles de variation due à des causes spéciales, aussi bien dans les moyennes des sous-groupes que dans la variation au sein des sous-groupes.",

// GRAPHIQUE CHRONOLOGIQUE — RÉFÉRENCE
"summary.run.baselineFirst": "Période de référence : les",
"summary.run.baselinePoints": "premiers points sont utilisés pour calculer la médiane.",
"summary.run.baselineAll": "Période de référence : tous les points de cette période sont utilisés pour calculer la médiane.",

// GRAPHIQUE CHRONOLOGIQUE — SIGNAUX
"summary.run.shiftBefore": "un décalage prolongé (≥",
"summary.run.shiftAfter": "points du même côté de la médiane)",
"summary.run.trendBefore": "une tendance prolongée (≥",
"summary.run.trendAfter": "points en augmentation ou en diminution)",
"summary.run.outlier": "une valeur exceptionnellement éloignée (point « astronomique »)",

// GRAPHIQUE CHRONOLOGIQUE — INTERPRÉTATION
"summary.run.noSignals": "Aucun signal clair de cause spéciale n'a été détecté pendant cette période (pas de décalage prolongé, de tendance prolongée ni de valeur exceptionnellement éloignée). Ce comportement est compatible avec une variation habituelle, mais doit toujours être interprété dans son contexte.",
"summary.run.hasSignals": "Cette période présente des signaux de cause spéciale :",

// GRAPHIQUE CHRONOLOGIQUE — POINTS À EXAMINER
"summary.run.where": "Points à examiner :",
"summary.run.whereShift": "décalage autour des points",
"summary.run.whereTrend": "tendance autour des points",
"summary.run.whereOutlierOne": "valeur exceptionnellement éloignée au point",
"summary.run.whereOutlierMany": "valeurs exceptionnellement éloignées aux points",
"summary.run.increasing": "croissante",
"summary.run.decreasing": "décroissante",

// GRAPHIQUE CHRONOLOGIQUE — CHANGEMENTS DE PHASE
"summary.run.splitNote": "Chaque période est résumée séparément, car les changements de phase suggèrent que le processus a pu évoluer au fil du temps.",

// XMR — DESCRIPTIONS DES SIGNAUX
"summary.xmr.signalBeyond": "un ou plusieurs points se situent au-delà des limites de contrôle",
"summary.xmr.signalShift": "un décalage prolongé (plusieurs points du même côté de la moyenne)",
"summary.xmr.signalTrend": "une tendance prolongée (augmentation ou diminution continue)",
"summary.xmr.signalOutlier": "une valeur exceptionnellement éloignée (point « astronomique »)",

// XMR — PÉRIODE DE RÉFÉRENCE
"summary.xmr.baselineWindowStart": "Période de référence : fenêtre initiale de",
"summary.xmr.baselineWindowOne": "point servant de référence ;",
"summary.xmr.baselineWindowMany": "points servant de référence ;",
"summary.xmr.baselineUsed": "pris en compte pour calculer la moyenne et les limites",
"summary.xmr.excludedOne": "point exclu du calcul",
"summary.xmr.excludedMany": "points exclus du calcul",
"summary.xmr.baselineAll": "Période de référence : tous les points de cette période ont été utilisés pour calculer la moyenne et les limites.",

// XMR — EXCLUSIONS MANUELLES
"summary.xmr.exclusionHeading": "Ajustement lié aux points exclus",
"summary.xmr.adjacentOne": "étendue mobile adjacente n'a pas été utilisée pour estimer la variation, car elle comprend un point exclu.",
"summary.xmr.adjacentMany": "étendues mobiles adjacentes n'ont pas été utilisées pour estimer la variation, car elles comprennent un point exclu.",
"summary.xmr.noArtificialMR": "SimpleSPC ne crée pas d'étendue mobile artificielle de part et d'autre d'un point exclu.",

// XMR — FILTRAGE DE NELSON
"summary.xmr.limitCalculation": "Calcul des limites XmR",
"summary.xmr.screenedOne": "variation exceptionnellement grande entre deux points consécutifs n'a pas été utilisée pour estimer la variation habituelle. Tous les points restent visibles sur la carte.",
"summary.xmr.screenedMany": "variations exceptionnellement grandes entre deux points consécutifs n'ont pas été utilisées pour estimer la variation habituelle. Tous les points restent visibles sur la carte.",
"summary.xmr.nelsonUsed": "Le filtrage des étendues mobiles selon la méthode de Nelson a été appliqué :",
"summary.xmr.initialMR": "moyenne initiale des étendues mobiles",
"summary.xmr.threshold": "seuil de filtrage",
"summary.xmr.revisedMR": "moyenne révisée des étendues mobiles",
"summary.xmr.screeningOff": "Le filtrage des étendues mobiles est désactivé ; toutes les variations entre points consécutifs de la période de référence ont donc été utilisées pour estimer la variation habituelle.",

// XMR — LIMITES PROVISOIRES
"summary.xmr.provisionalHeading": "Limites provisoires",
"summary.xmr.provisionalOnly": "seulement",
"summary.xmr.provisionalOne": "point a été utilisé pour estimer les limites de cette période.",
"summary.xmr.provisionalMany": "points ont été utilisés pour estimer les limites de cette période.",
"summary.xmr.minimumIntro": "SimpleSPC nécessite au moins",
"summary.xmr.minimumEnd": "points avant de pouvoir interpréter les signaux de cause spéciale ou estimer la capabilité du processus.",
"summary.xmr.provisionalContext": "Les limites affichées sont fournies uniquement à titre indicatif.",

// XMR — OBJECTIFS
"summary.xmr.target": "Objectif",
"summary.xmr.aboveBetter": "une valeur supérieure ou égale est préférable",
"summary.xmr.belowBetter": "une valeur inférieure ou égale est préférable",
"summary.xmr.targetCoverage": "des points de cette période atteignent l'objectif",
"summary.xmr.coverageUnavailable": "La proportion de points atteignant l'objectif n'a pas été calculée pour cette période.",

// XMR — INTERPRÉTATION
"summary.xmr.notAssessed": "Évaluation non encore possible : davantage de données sont nécessaires avant de pouvoir déterminer si cette période présente une variation habituelle ou des signaux de cause spéciale.",
"summary.xmr.noSignals": "Aucun signal clair de cause spéciale n'a été détecté pendant cette période. Le comportement observé est compatible avec une variation habituelle, mais doit toujours être interprété dans son contexte clinique.",
"summary.xmr.hasSignals": "Cette période présente des signaux de cause spéciale :",
"summary.xmr.where": "Points à examiner",
"summary.xmr.outsideOne": "point au-delà des limites au rang",
"summary.xmr.outsideMany": "points au-delà des limites aux rangs",
"summary.xmr.outlierOne": "valeur exceptionnellement éloignée au rang",
"summary.xmr.outlierMany": "valeurs exceptionnellement éloignées aux rangs",

// XMR — CAPABILITÉ
"summary.xmr.estimatedCapability": "Capabilité estimée (cette période)",
"summary.xmr.ifStable": "si le processus reste stable, environ",
"summary.xmr.futureMeet": "des points futurs devraient atteindre l'objectif.",
"summary.xmr.capability": "Capabilité",
"summary.xmr.capabilityUnreliable": "un objectif a été défini, mais des signaux de cause spéciale sont présents pendant cette période ; toute estimation de capabilité serait donc peu fiable.",
"summary.xmr.comparePeriods": "La comparaison des moyennes, des limites et des performances par rapport à l'objectif entre les périodes peut aider à déterminer si le processus a changé après des interventions.",

// XMR — INDICATEUR DE CAPABILITÉ
"summary.xmr.badgeHeading": "CAPABILITÉ DU PROCESSUS (dernière période)",
"summary.xmr.badgeNote": "(Probabilité estimée d'atteindre l'objectif pendant la dernière période, sous l'hypothèse d'un processus stable et d'une distribution approximativement normale.)",
"summary.xmr.badgeWarning": "Signaux de cause spéciale détectés pendant la dernière période :",
"summary.xmr.badgeUnreliable": "les estimations de capabilité peuvent être peu fiables tant que ces signaux ne sont pas compris.",
"summary.xmr.badgeAction": "Cherchez à comprendre et à traiter ces causes avant de vous appuyer sur des estimations de capabilité.",

"summary.exclusions.heading": "Exclus du calcul des limites :",
"summary.exclusions.point": "point",
"summary.exclusions.points": "points",
"summary.exclusions.subgroup": "sous-groupe",
"summary.exclusions.subgroups": "sous-groupes",
"summary.exclusions.pointVisible": "Le point exclu reste visible sur le graphique.",
"summary.exclusions.pointsVisible": "Les points exclus restent visibles sur le graphique.",
"summary.exclusions.subgroupVisible": "Le sous-groupe exclu reste visible sur le graphique.",
"summary.exclusions.subgroupsVisible": "Les sous-groupes exclus restent visibles sur le graphique.",

// CARTES AUX ATTRIBUTS — PÉRIODE DE RÉFÉRENCE
"summary.attr.baselineTitle": "Période de référence pour cette période",
"summary.attr.firstOne": "le premier point (1) définit la fenêtre de référence ;",
"summary.attr.firstBefore": "les",
"summary.attr.firstAfter": "premiers points définissent la fenêtre de référence ;",
"summary.attr.used": "utilisés pour calculer la ligne centrale et les limites",
"summary.attr.excludedOne": "exclu du calcul",
"summary.attr.excludedMany": "exclus du calcul",

// CARTES AUX ATTRIBUTS — SIGNAUX ET INTERPRÉTATION
"summary.attr.signals": "Signaux",
"summary.attr.stable": "Selon les règles sélectionnées, aucun signal clair de cause spéciale n'a été détecté pendant cette période. Cela suggère un comportement compatible avec une variation habituelle (de cause commune).",
"summary.attr.unstable": "Selon les règles sélectionnées, des signaux de cause spéciale ont été détectés pendant cette période (un comportement peu susceptible de s'expliquer uniquement par la variation habituelle).",

// CARTES AUX ATTRIBUTS — PRÉCAUTIONS STATISTIQUES
"summary.attr.caution": "Ces règles sont des indications, et non des réponses définitives. Interprétez les résultats en tenant compte du contexte local (modifications du processus, effectifs, demande, définitions ou codage) et des hypothèses de base des cartes SPC, notamment la comparabilité des mesures et des opportunités d'observation au fil du temps.",

// CARTES AUX ATTRIBUTS — SIGNAUX INDIVIDUELS
"summary.attr.signalAbove": "Un ou plusieurs points dépassent la limite supérieure de contrôle",
"summary.attr.signalBelow": "Un ou plusieurs points se situent sous la limite inférieure de contrôle",

"summary.attr.shiftName": "Décalage",
"summary.attr.shiftAbove": "points consécutifs au-dessus de la ligne centrale (de",
"summary.attr.shiftBelow": "points consécutifs au-dessous de la ligne centrale (de",

"summary.attr.trendName": "Tendance",
"summary.attr.trendIncreasing": "points en augmentation régulière (de",
"summary.attr.trendDecreasing": "points en diminution régulière (de",

"summary.attr.zone23": "Zone : 2 points sur 3 dans le tiers extérieur (de",
"summary.attr.zone45": "Zone : 4 points sur 5 au-delà de 1 sigma (de",

  // DIAGNOSTICS LANEY P / P'
  "summary.laney.p.suggest": "La carte P′ pourrait être mieux adaptée à ces données.",
  "summary.laney.p.over": "Les proportions varient davantage d'un point à l'autre que ne le prévoit une carte P standard pour ces tailles d'échantillon. Les limites d'une carte P standard peuvent donc être trop étroites et générer des alertes supplémentaires. Une carte P′ ajuste les limites pour tenir compte de cette variation supplémentaire, tout en conservant les mêmes proportions et la même ligne centrale.",
  "summary.laney.p.under": "Les proportions varient moins d'un point à l'autre que ne le prévoit une carte P standard. Les limites de contrôle standard peuvent donc être trop larges et rendre les changements réels plus difficiles à détecter. Une carte P′ ajuste les limites pour mieux refléter la variation effectivement observée.",
  "summary.laney.p.normalTitle": "Une carte P standard semble adaptée à ces données.",
  "summary.laney.p.normal": "La variation observée se situe dans la plage attendue pour des données de proportions avec ces tailles d'échantillon. Il n'y a donc pas de raison évidente de privilégier une carte P′.",

  "summary.laney.p.prime.over": "Ces données varient davantage que ne le prévoit une carte P standard ; la carte P′ utilise donc des limites de contrôle plus larges.",
  "summary.laney.p.prime.under": "Ces données varient moins que ne le prévoit une carte P standard ; la carte P′ utilise donc des limites de contrôle plus étroites.",
  "summary.laney.p.prime.normal": "La variation observée est très proche de celle attendue pour une carte P standard ; l'ajustement est donc faible.",

  // DIAGNOSTICS LANEY U / U'
  "summary.laney.u.suggest": "La carte U′ pourrait être mieux adaptée à ces données.",
  "summary.laney.u.over": "Les taux varient davantage d'un point à l'autre que ne le prévoit une carte U standard, compte tenu du nombre d'opportunités. Les limites d'une carte U standard peuvent donc être trop étroites et générer des alertes supplémentaires. Une carte U′ ajuste les limites pour tenir compte de cette variation supplémentaire, tout en conservant les mêmes taux et la même ligne centrale.",
  "summary.laney.u.under": "Les taux varient moins d'un point à l'autre que ne le prévoit une carte U standard. Les limites de contrôle standard peuvent donc être trop larges et rendre les changements réels plus difficiles à détecter. Une carte U′ ajuste les limites pour mieux refléter la variation effectivement observée.",
  "summary.laney.u.normalTitle": "Une carte U standard semble adaptée à ces données.",
  "summary.laney.u.normal": "La variation observée se situe dans la plage attendue pour des données de taux de non-conformités. Il n'y a donc pas de raison évidente de privilégier une carte U′.",

  "summary.laney.u.prime.over": "Ces données varient davantage que ne le prévoit une carte U standard ; la carte U′ utilise donc des limites de contrôle plus larges.",
  "summary.laney.u.prime.under": "Ces données varient moins que ne le prévoit une carte U standard ; la carte U′ utilise donc des limites de contrôle plus étroites.",
  "summary.laney.u.prime.normal": "La variation observée est très proche de celle attendue pour une carte U standard ; l'ajustement est donc faible.",

  // DIAGNOSTICS DE DISPERSION COMMUNS
  "summary.laney.technicalCheck": "Vérification technique",
  "summary.laney.observedVariation": "variation observée",
  "summary.laney.ofExpected": "de la variation attendue",
  "summary.laney.diagnosticUpper": "limite diagnostique supérieure",
  "summary.laney.diagnosticLower": "limite diagnostique inférieure",
  "summary.laney.expectedRange": "plage attendue",
  "summary.laney.switchP": "Utiliser plutôt la carte P′",
  "summary.laney.switchU": "Utiliser plutôt la carte U′",
  "summary.laney.whyDifferent": "Pourquoi ces limites sont-elles différentes ?",
  "summary.laney.adjustmentBenefit": "Cet ajustement permet aux limites de contrôle de mieux refléter la variation effectivement observée dans les données.",
"summary.laney.technicalDetail": "Détail technique",

// RÉSUMÉS DES CARTES T ET G
"summary.rare.baseline": "Points initiaux de cette période utilisés pour calculer la ligne centrale et les limites",

"summary.rare.stable": "Aucun signal clair de cause spéciale n'a été détecté pendant cette période. Le comportement observé est compatible avec une variation habituelle (de cause commune).",
"summary.rare.unstable": "Des signaux de cause spéciale ont été détectés pendant cette période (un comportement incompatible avec la seule variation habituelle).",

"summary.rare.betterHeading": "Interpréter une amélioration",
"summary.rare.better": "Si l'événement est indésirable, des intervalles plus longs entre les événements (ou davantage d'opportunités entre eux) sont généralement préférables. Si l'événement est souhaitable, des intervalles plus courts peuvent être préférables.",

"summary.rare.nextHeading": "Que faire ensuite",
"summary.rare.nextStable": "Si le niveau de performance est insuffisant, privilégiez l'amélioration du processus plutôt que de réagir à chaque point individuel.",
"summary.rare.nextUnstable": "Recherchez une explication concrète (changement de processus, effectifs, demande, détection ou définition des événements). Si le changement était planifié, envisagez une nouvelle période de référence une fois la situation stabilisée.",

"summary.rare.dataHeading": "Rappel sur les données",
"summary.rare.dataT": "Cette carte utilise le temps écoulé entre les événements (par exemple, le nombre de jours entre deux incidents).",
"summary.rare.dataG": "Cette carte utilise le nombre d'opportunités entre les événements (les valeurs doivent être supérieures ou égales à 1).",

// MENU CONTEXTUEL — BOUTONS
"context.annotations": "Annotations ▸",
"context.annotationAdd": "Ajouter une annotation…",
"context.annotationEdit": "Modifier une annotation…",
"context.annotationDelete": "Supprimer une annotation…",
"context.annotationClear": "Supprimer toutes les annotations",

"context.excludePoint": "Exclure du calcul des limites…",
"context.excludeSubgroup": "Exclure ce sous-groupe du calcul des limites…",
"context.includePoint": "Réintégrer dans le calcul des limites",
"context.includeSubgroup": "Réintégrer ce sous-groupe dans le calcul des limites",

"context.splits": "Changements de phase ▸",
"context.addSplit": "Ajouter un changement de phase après ce point…",
"context.removeSplit": "Supprimer le changement de phase après {label} (point {point})",
"context.clearSplits": "Supprimer tous les changements de phase",
"context.noSplits": "Aucun changement de phase à supprimer",
"context.splitsUnavailable": "Changements de phase indisponibles pour ce type de carte",

"context.centreShow": "Afficher les libellés des lignes centrales",
"context.centreHide": "Masquer les libellés des lignes centrales",
"context.centreHideOne": "Masquer ce libellé de ligne centrale",
"context.centreRestore": "Rétablir les libellés de ligne centrale masqués",

"context.copy": "Copier les graphiques",
"context.copyAnalysis": "Copier les graphiques et l'analyse",
"context.savePng": "Enregistrer les graphiques en PNG",
"context.pdf": "Télécharger le rapport PDF",

// MENU CONTEXTUEL — INFOBULLES
"context.tip.annotations": "Ajouter, modifier ou supprimer des annotations pour le point sélectionné.",
"context.tip.annotationAdd": "Ajouter une annotation à ce point.",
"context.tip.annotationEdit": "Modifier une annotation existante à ce point.",
"context.tip.annotationDelete": "Supprimer une annotation existante à ce point.",
"context.tip.annotationClear": "Supprimer toutes les annotations du graphique.",

"context.tip.excludeDefault": "Conserver ce point visible, mais l'exclure de l'estimation de la ligne centrale et des limites de contrôle.",
"context.tip.includeDefault": "Réintégrer ce point dans le calcul de la ligne centrale et des limites de contrôle.",
"context.tip.excludeEligible": "Conserver ce point visible sans l'utiliser pour estimer la ligne centrale ou les limites de contrôle.",
"context.tip.nearPoint": "Faites un clic droit près d'un point du graphique.",
"context.tip.outsideBaseline": "Ce point se trouve en dehors de la période de référence utilisée pour estimer les limites actuelles. Son exclusion ne modifierait donc pas leur calcul.",
"context.tip.tooFewPoint": "L'exclusion de ce point laisserait trop peu d'observations incluses pour estimer les limites de contrôle.",
"context.tip.tooFewSubgroup": "L'exclusion de ce sous-groupe laisserait trop peu d'observations incluses pour estimer les limites de contrôle.",
"context.tip.noConsecutive": "L'exclusion de ce point ne laisserait aucune paire d'observations consécutives incluses pour estimer la variation à partir des étendues mobiles.",

"context.tip.splits": "Ajouter ou supprimer des changements de phase sur cette carte.",
"context.tip.addSplit": "Les changements de phase entraînent un nouveau calcul de la ligne centrale et des limites de contrôle à partir du point choisi. Ils ne doivent être utilisés que lorsqu'un changement durable de performance a été identifié, étudié, compris et reconnu comme représentatif du nouveau niveau habituel du processus.",
"context.tip.splitsUnavailable": "Les changements de phase ne sont pas disponibles pour ce type de carte.",
"context.tip.splitPoint": "Faites un clic droit près d'un point pour ajouter un changement de phase.",
"context.tip.removeSplit": "Supprimer le changement de phase après {label}.",
"context.tip.clearSplits": "Supprimer tous les changements de phase et revenir à un ensemble unique de limites.",
"context.tip.noSplits": "Aucun changement de phase à supprimer.",

"context.tip.centreLabels": "Afficher ou masquer les valeurs moyennes calculées pour cette carte XmR.",
"context.tip.centreHideOne": "Masquer uniquement le libellé de moyenne sur lequel vous avez fait un clic droit.",
"context.tip.centreRestore": "Rétablir les libellés de moyenne précédemment masqués.",

"context.tip.copy": "Copier uniquement le graphique dans le presse-papiers.",
"context.tip.copyAnalysis": "Copier le graphique et la zone d'analyse dans le presse-papiers.",
"context.tip.savePng": "Enregistrer une copie du graphique au format PNG.",
"context.tip.pdf": "Exporter le graphique et l'analyse au format PDF.",

// OPTIONS D'ANNOTATION DÉSACTIVÉES
"context.tip.annotationPointAdd": "Faites un clic droit près d'un point pour ajouter une annotation.",
"context.tip.annotationPointEdit": "Faites un clic droit près d'un point pour modifier une annotation.",
"context.tip.annotationMissingEdit": "Aucune annotation ne peut être modifiée à ce point.",
"context.tip.annotationPointDelete": "Faites un clic droit près d'un point pour supprimer une annotation.",
"context.tip.annotationMissingDelete": "Aucune annotation ne peut être supprimée à ce point.",
"context.tip.annotationNone": "Aucune annotation à supprimer.",

// DIALOGUES DES ACTIONS SUR LES CARTES
"dialog.noChart": "Veuillez d'abord générer un graphique.",
"dialog.nearPoint": "Faites d'abord un clic droit près d'un point du graphique.",
"dialog.notInWindow": "Ce point ne fait pas partie de la période utilisée pour le calcul des limites actuelles.",

// EXCLUSION DES OBSERVATIONS
"dialog.excludePoint": "Exclure ce point du calcul des limites ?\n\nLe point restera visible sur le graphique, mais il ne sera pas utilisé pour estimer la ligne centrale ou les limites de contrôle.\n\nNe procédez à cette exclusion qu'après avoir étudié ce point et si vous avez de bonnes raisons de penser qu'il correspond à une cause spéciale identifiée plutôt qu'à la variation habituelle du processus.\n\nN'excluez pas un point simplement parce que sa valeur est exceptionnellement élevée ou faible. S'il représente un changement durable du processus, utilisez plutôt un changement de phase.",
"dialog.excludeSubgroup": "Exclure ce sous-groupe du calcul des limites ?\n\nLe sous-groupe restera visible sur le graphique, mais il ne sera pas utilisé pour estimer la ligne centrale ou les limites de contrôle.\n\nNe procédez à cette exclusion qu'après avoir étudié ce sous-groupe et si vous avez de bonnes raisons de penser qu'il correspond à une cause spéciale identifiée plutôt qu'à la variation habituelle du processus.\n\nN'excluez pas un point simplement parce que sa valeur est exceptionnellement élevée ou faible. S'il représente un changement durable du processus, utilisez plutôt un changement de phase.",
"dialog.excludedPointNote": "Le point a été exclu du calcul des limites. Il restera visible sur le graphique.\n\nSouhaitez-vous ajouter une annotation pour expliquer cette exclusion ?",
"dialog.excludedSubgroupNote": "Le sous-groupe a été exclu du calcul des limites. Il restera visible sur le graphique.\n\nSouhaitez-vous ajouter une annotation pour expliquer cette exclusion ?",


// ANNOTATIONS
"dialog.annotationFor": "Annotation pour {label} :",
"dialog.newAnnotationFor": "Nouvelle annotation pour {label} :",
"dialog.editAnnotationFor": "Modifier l'annotation pour {label} :",
"dialog.annotationPositionMissing": "Impossible de déterminer la position sélectionnée sur l'axe horizontal pour l'annotation.",
"dialog.annotationBlank": "Le texte de l'annotation ne peut pas être vide.",
"dialog.editWhich": "Quelle annotation souhaitez-vous modifier ?",
"dialog.deleteWhich": "Quelle annotation souhaitez-vous supprimer ?",
"dialog.annotationsAt": "Annotations pour {label} :",
"dialog.chooseNumber": "Saisissez un numéro entre 1 et {count} :",
"dialog.invalidNumber": "Veuillez saisir un numéro valide.",
"dialog.noEditAt": "Aucune annotation à modifier pour {label}.",
"dialog.noDeleteAt": "Aucune annotation à supprimer pour {label}.",
"dialog.deleteAnnotation": "Supprimer cette annotation ?\n\n« {label} »",
"dialog.clearAnnotations": "Supprimer toutes les annotations ?",

// VALIDATION DES CHANGEMENTS DE PHASE
"dialog.splitUnavailable": "Les changements de phase et le recalcul des limites ne sont pas disponibles pour ce type de carte.",
"dialog.splitInvalid": "Veuillez sélectionner un point de changement de phase valide.",
"dialog.splitLast": "Il n'est pas possible d'ajouter un changement de phase après le dernier point.",

// COPIE ET EXPORTATION
"dialog.copyDone": "L'image du graphique a été copiée dans le presse-papiers.",
"dialog.copyAnalysisDone": "L'image du graphique et de l'analyse a été copiée dans le presse-papiers.",
"dialog.actionFailed": "Cette action a échoué dans ce navigateur. Essayez plutôt d'enregistrer les graphiques au format PNG.",
"dialog.pdfImageFailed": "Impossible de créer l'image destinée au rapport PDF.",
"dialog.pdfLibraryMissing": "La bibliothèque d'exportation PDF n'est pas disponible.",

// ÉDITEUR DE DONNÉES
"editor.title": "Éditeur de données",
"editor.intro": "Collez directement les données dans la grille, ou utilisez cet éditeur pour examiner les feuilles Excel avant l'importation. Vous pouvez changer de feuille, supprimer les lignes ou colonnes inutiles, puis cliquer sur",
"editor.helpButton": "Comment supprimer des lignes ou des colonnes ?",
"editor.helpAria": "Comment supprimer des lignes ou des colonnes ?",
"editor.helpHeading": "Conseil de modification",
"editor.helpRow": "Faites un clic droit sur un numéro de ligne pour supprimer une ligne",
"editor.helpColumn": "Faites un clic droit sur un en-tête de colonne pour supprimer une colonne",
"editor.helpInsert": "Vous pouvez aussi insérer des lignes ou des colonnes de la même manière",
"editor.worksheet": "Feuille :",
"editor.hasHeaders": "La première ligne contient les en-têtes",
"editor.detect": "Détection automatique",
"editor.apply": "Appliquer les données",
"editor.cancel": "Annuler",
"editor.chooseWorksheet": "Choisir une feuille",
"editor.chooseDescription": "Ce fichier Excel contient plusieurs feuilles. Choisissez celle que vous souhaitez importer.",
"editor.loadSheet": "Charger la feuille",

// MESSAGES DYNAMIQUES DE L'ÉDITEUR
"editor.workbookStatus": "Classeur chargé. Vous pouvez changer de feuille ici, supprimer les lignes ou colonnes inutiles dans la grille, puis cliquer sur <strong>Appliquer les données</strong>.",
"editor.headerApplyHeadings": "Lors de l'application, la <strong>première ligne de la grille</strong> sera traitée comme des <strong>en-têtes</strong>.<br><small>Conseil : si vous avez importé un fichier CSV normalement, les en-têtes sont déjà les titres des colonnes. Laissez donc généralement cette option désactivée.</small>",
"editor.headerApplyData": "Lors de l'application, la <strong>première ligne de la grille</strong> sera traitée comme des <strong>données</strong>.<br><small>Conseil : si vous avez importé un fichier CSV normalement, les en-têtes sont déjà les titres des colonnes. Laissez donc généralement cette option désactivée.</small>",
"editor.detectHeadings": "Détection automatique : la première ligne semble contenir des <strong>en-têtes</strong>.",
"editor.detectData": "Détection automatique : la première ligne semble contenir des <strong>données</strong>.",
"editor.headerMismatchChecked": "Vous avez coché « La première ligne contient les en-têtes », mais la première ligne semble contenir des DONNÉES.\n\nAppliquer malgré tout ?",
"editor.headerMismatchUnchecked": "Vous n'avez pas coché « La première ligne contient les en-têtes », mais la première ligne semble contenir des EN-TÊTES.\n\nAppliquer malgré tout ?",
"editor.notInitialised": "L'éditeur de tableau n'a pas été initialisé. Essayez de le rouvrir.",
"editor.enterRow": "Veuillez saisir au moins une ligne de données.",
"editor.pasteRow": "Collez au moins une ligne de données.",
"editor.readError": "Une erreur inattendue s'est produite lors de la lecture des données du tableau.",

// RÉINITIALISATION ET REMPLACEMENT
"data.confirmReset": "Réinitialiser ce graphique ?\n\nLes données chargées et les colonnes X/valeur sélectionnées seront conservées, mais les paramètres du graphique (type de carte, période de référence, objectif, changements de phase, annotations, libellés, règles et mise en forme des axes) seront réinitialisés.",
"data.confirmClear": "Effacer toutes les données chargées ?\n\nCette action réinitialisera également le graphique actuel et ses paramètres. Enregistrez le graphique au préalable si vous souhaitez le rouvrir plus tard.",
"data.confirmReplace": "Charger ces données comme un nouveau graphique ?\n\nLes données et les paramètres du graphique actuel seront effacés. Enregistrez d'abord le graphique actuel si vous souhaitez le rouvrir plus tard.",
"data.noRows": "Aucune ligne de données n'a été trouvée.",
"data.noColumns": "Aucune colonne n'a pu être détectée dans les données.",
"data.fileType": "Veuillez importer un fichier CSV ou Excel (.csv, .xlsx, .xls).",
"data.csvError": "Erreur lors de l'analyse du fichier CSV : {message}",
"data.csvNoHeaders": "Il semble que votre fichier CSV ne contienne pas d'en-têtes de colonnes.\n\nCliquez sur OK pour traiter la première ligne comme des DONNÉES (les colonnes seront nommées Column1, Column2, etc.).\nCliquez sur Annuler si la première ligne contient des en-têtes (ajoutez les en-têtes, puis importez à nouveau le fichier).",
"data.csvAddHeaders": "Ajoutez une ligne d'en-têtes (par exemple Date,Value), puis importez à nouveau le fichier.",
"data.excelReadError": "Une erreur inattendue s'est produite lors de la lecture du fichier Excel.",
"data.csvReadError": "Une erreur inattendue s'est produite lors de la lecture du fichier CSV.",

// PROJETS JSON ET ANCIENS PARAMÈTRES
"project.invalidObject": "Ce fichier n'a pas pu être interprété comme un projet SPC.",
"project.noSavedData": "Ce fichier de projet ne contient aucune donnée enregistrée.",
"project.loadFailed": "Les données du projet n'ont pas pu être chargées.",
"project.settingsQueued": "Paramètres chargés. Importez maintenant vos données CSV ou Excel ; l'outil appliquera automatiquement ces paramètres.",
"project.unsupported": "Ce fichier ne semble pas être un projet SPC ou un fichier de paramètres pris en charge.",
"project.invalidJson": "Impossible de lire ce fichier JSON. Vérifiez qu'il s'agit d'un export valide de cet outil.",
"project.restoreFailed": "Les données existantes n'ont pas pu être restaurées après la réinitialisation du graphique.",

// VALIDATION AVANT GÉNÉRATION DU GRAPHIQUE
"validation.selectColumns": "Veuillez sélectionner une colonne pour l'axe X et une colonne de valeurs.",

//  AIDE ET CONSEILS : SECTIONS 1 À 3

// Fenêtre d'aide
"help.subtitle": "Guide pratique pour importer des données, créer des cartes, interpréter les signaux SPC et enregistrer votre travail.",
"help.closeAria": "Fermer l'aide",

// 1. Présentation de l'outil
"help.1.title": "1. À quoi sert cet outil",
"help.1.intro": "est un outil gratuit, accessible dans un navigateur, permettant de créer des cartes de maîtrise statistique des procédés (SPC) et des diagrammes de Pareto.",
"help.1.purpose": "Il vous aide à comprendre les tendances et les comportements observés dans les données, à distinguer la variation habituelle des signaux de changement et à identifier des possibilités d'amélioration. Tous les calculs sont effectués localement dans votre navigateur. Vos données ne sont pas envoyées à un serveur.",
"help.1.scope": "SimpleSPC est conçu pour l'apprentissage, l'amélioration de la qualité, le suivi des services et les analyses pratiques. Il n'a pas vocation à remplacer un logiciel statistique réglementé.",
"help.1.available": "Cartes et graphiques disponibles :",

"help.chart.runName": "Carte de suivi (Run)",
"help.chart.runDesc": "présente les valeurs dans l'ordre chronologique ou séquentiel, avec une médiane et des règles simples permettant d'identifier des comportements non aléatoires.",
"help.chart.xmrName": "Carte XmR",
"help.chart.xmrDesc": "présente les mesures individuelles avec une moyenne et des limites de contrôle, ainsi qu'une carte des étendues mobiles facultative.",
"help.chart.xbarsName": "Carte X̄–S",
"help.chart.xbarsDesc": "examine les moyennes et la variation à l'intérieur de sous-groupes de mesures continues.",
"help.chart.cName": "Carte C",
"help.chart.cDesc": "surveille les nombres d'événements ou de défauts lorsque les opportunités de survenue restent relativement constantes.",
"help.chart.pName": "Cartes P et P′ de Laney",
"help.chart.pDesc": "surveillent les proportions, par exemple le pourcentage de patients concernés par un événement. La carte P′ ajuste les limites de contrôle lorsque la variation observée diffère de celle attendue avec une carte P standard.",
"help.chart.uName": "Cartes U et U′ de Laney",
"help.chart.uDesc": "surveillent les taux d'événements lorsque le nombre d'opportunités varie. La carte U′ ajuste les limites de contrôle lorsque la variation observée diffère de celle attendue avec une carte U standard.",
"help.chart.tName": "Carte T",
"help.chart.tDesc": "surveille le temps écoulé entre des événements rares.",
"help.chart.gName": "Carte G",
"help.chart.gDesc": "surveille le nombre d'opportunités entre des événements rares.",
"help.chart.paretoName": "Diagramme de Pareto",
"help.chart.paretoDesc": "classe les catégories de la plus fréquente à la moins fréquente afin d'aider à prioriser les investigations et les améliorations.",

"help.1.importantTitle": "Important :",
"help.1.important": "Les diagrammes de Pareto sont différents des cartes SPC ordonnées dans le temps. Ils permettent de déterminer quelles catégories représentent le plus grand nombre d'occurrences, mais ne permettent pas d'établir si un processus est stable ou si un changement est statistiquement significatif.",

// 2. Démarrage rapide
"help.2.title": "2. Démarrage rapide",
"help.2.step1": "Importez des données au format CSV ou Excel, ou utilisez l'éditeur de données.",
"help.2.step2": "Sélectionnez le type de carte et choisissez les colonnes appropriées.",
"help.2.step3Before": "Cliquez sur",
"help.2.step3After": "si le graphique ne se met pas à jour automatiquement.",
"help.2.step4": "Ajoutez des titres, des libellés, des objectifs, des changements de phase ou des annotations si nécessaire.",
"help.2.step5": "Examinez le résumé et les signaux indiqués sur le graphique.",
"help.2.step6": "Enregistrez le graphique si vous souhaitez reprendre votre travail ultérieurement.",

// 3. Préparation et importation des données
"help.3.title": "3. Préparer et importer vos données",
"help.3.intro": "Les colonnes nécessaires dépendent du type de carte. La plupart des cartes SPC nécessitent une colonne de dates ou de séquence et une mesure ou un effectif.",

"help.3.pName": "Les cartes P, P′, U et U′",
"help.3.pDesc": "nécessitent également une colonne de dénominateurs ou d'opportunités.",
"help.3.xbarsName": "Les cartes X̄–S",
"help.3.xbarsDesc": "nécessitent un identifiant de sous-groupe et les valeurs des mesures individuelles.",
"help.3.paretoName": "Les diagrammes de Pareto",
"help.3.paretoDesc": "nécessitent une colonne de catégories et, si vos données sont déjà regroupées, éventuellement une colonne de fréquences ou d'effectifs.",

"help.3.ways": "Vous pouvez importer des données de trois façons :",
"help.3.csv": "Importer un fichier CSV",
"help.3.excel": "Importer un classeur Excel",
"help.3.paste": "Coller ou modifier directement des données",
"help.3.pasteAfter": "dans l'éditeur de données",

"help.3.excelBefore": "Si vous importez un classeur Excel, l'outil ouvre",
"help.3.excelMiddle": "pour vous permettre d'examiner la feuille avant de l'importer. Vous pouvez changer de feuille, supprimer les lignes ou colonnes inutiles, puis cliquer sur",
"help.3.axisAuto": "L'outil essaie de proposer automatiquement des colonnes X et Y adaptées, mais vous pouvez les modifier à tout moment.",
"help.3.dateBefore": "Si votre colonne de dates utilise des dates numériques telles que",
"help.3.dateMiddle": "utilisez le paramètre",
"help.3.dateSetting": "Format des dates numériques",
"help.3.dateAfter": "afin que l'outil les interprète correctement.",
"help.3.default": "Par défaut :",
"help.3.defaultFormat": "format britannique jour en premier (jj/mm/aaaa)",

// AIDE ET CONSEILS : SECTIONS 4 À 6

// 4. Utilisation de l'éditeur de données
"help.4.title": "4. Utiliser l'éditeur de données",
"help.4.intro": "L'éditeur de données est utile lorsque :",
"help.4.paste": "vous souhaitez coller des données provenant d'Excel",
"help.4.sheets": "votre classeur Excel contient plusieurs feuilles",
"help.4.trim": "votre feuille contient des lignes ou des colonnes inutiles à supprimer",
"help.4.sheetBefore": "Utilisez le sélecteur",
"help.4.sheetName": "Feuille",
"help.4.sheetAfter": "pour passer d'une feuille à l'autre dans un fichier Excel importé.",
"help.4.headersBefore": "Cochez ou décochez l'option",
"help.4.headersAfter": "selon que la première ligne contient des en-têtes ou des données.",
"help.4.detectBefore": "Utilisez la",
"help.4.detectAfter": "comme suggestion rapide, mais vérifiez toujours que le résultat est cohérent.",
"help.4.rowAction": "Faites un clic droit sur un numéro de ligne",
"help.4.rowAfter": "pour supprimer ou insérer des lignes.",
"help.4.columnAction": "Faites un clic droit sur un en-tête de colonne",
"help.4.columnAfter": "pour supprimer ou insérer des colonnes.",
"help.4.multiple": "Vous pouvez sélectionner plusieurs lignes ou colonnes avant de les supprimer.",
"help.4.tip": "Conseil : l'éditeur de données permet de nettoyer une feuille directement dans l'outil avant d'appliquer les données au graphique.",

// 5. Type d'axe X et type de carte
"help.5.title": "5. Type d'axe X et type de carte",
"help.5.intro": "Choisissez le type de carte adapté à la nature de vos données. Les principales possibilités sont :",
"help.5.run": "un point de départ simple pour représenter des valeurs au fil du temps.",
"help.5.xmr": "des mesures individuelles continues, avec des limites de contrôle.",
"help.5.xbars": "des mesures continues regroupées en sous-groupes.",
"help.5.c": "des nombres d'événements lorsque les opportunités de survenue restent relativement constantes.",
"help.5.p": "des proportions calculées à partir d'un numérateur et d'un dénominateur. La carte P′ tient compte d'une variation plus forte ou plus faible que celle attendue.",
"help.5.u": "des taux d'événements avec un nombre variable d'opportunités. La carte U′ tient compte d'une variation plus forte ou plus faible que celle attendue.",
"help.5.tgName": "Cartes T et G",
"help.5.tg": "le temps ou le nombre d'opportunités entre des événements rares.",
"help.5.pareto": "classe les catégories de la plus fréquente à la moins fréquente afin de prioriser les investigations.",
"help.5.xaxis": "Type d'axe X :",
"help.5.dateName": "Date / heure",
"help.5.date": "à utiliser lorsque les observations comportent des dates ou des heures et doivent être présentées dans l'ordre chronologique.",
"help.5.sequenceName": "Séquence / catégorie",
"help.5.sequence": "à utiliser lorsque les observations doivent être présentées dans l'ordre, sans interpréter leurs libellés comme des dates.",
"help.5.paretoNote": "Les diagrammes de Pareto utilisent des catégories plutôt que des dates. Les catégories sont classées par fréquence, avec une courbe de pourcentage cumulé. Ce ne sont pas des cartes SPC ordonnées dans le temps et ils ne comportent pas de limites de contrôle statistiques.",
"help.5.chooseBefore": "Si vous ne savez pas quelle carte utiliser, cliquez sur",
"help.5.chooseMiddle": "Vous pouvez également consulter",
"help.5.chooseAfter": "pour obtenir des conseils sur les colonnes nécessaires à la carte sélectionnée.",

// 6. Référence, phases, objectifs et menu contextuel
"help.6.title": "6. Périodes de référence, changements de phase, objectifs et menu contextuel",
"help.6.intro": "SimpleSPC propose des fonctions supplémentaires pour analyser vos données et expliquer les résultats.",
"help.6.baselineHeading": "Définir une période de référence :",
"help.6.baselineBefore": "Ouvrez l'onglet",
"help.6.baselineAfter": "et indiquez le nombre de premiers points à utiliser comme période de référence. Laissez ce champ vide si vous n'avez pas besoin d'une période de référence initiale distincte.",
"help.6.targetHeading": "Définir un objectif :",
"help.6.targetBefore": "Ouvrez l'onglet",
"help.6.targetMiddle": "et recherchez le paramètre",
"help.6.targetName": "Objectif",
"help.6.targetAfter": "Saisissez le niveau de performance souhaité, puis indiquez si les valeurs supérieures ou égales, ou inférieures ou égales, à cet objectif sont favorables. Vous pouvez également afficher ou masquer la ligne d'objectif. Un objectif est un niveau de performance visé, et non une limite de contrôle statistique.",
"help.6.splitHeading": "Ajouter un changement de phase :",
"help.6.splitBefore": "Faites un clic droit sur un point du graphique et sélectionnez",
"help.6.splitsMenu": "Changements de phase",
"help.6.splitMiddle": "puis",
"help.6.addSplitAction": "Ajouter un changement de phase après ce point",
"help.6.splitAfter": "Cette action divise le graphique en périodes d'analyse distinctes lorsque cette fonction est disponible.",
"help.6.annotationHeading": "Ajouter une annotation :",
"help.6.annotationBefore": "Faites un clic droit sur un point du graphique et utilisez le menu",
"help.6.annotationsMenu": "Annotations",
"help.6.annotationAfter": "pour consigner un événement, une intervention ou une explication.",

"help.6.menuHeading": "Découvrir le menu contextuel",
"help.6.menuIntro": "Faites un clic droit directement sur le graphique pour accéder à des options supplémentaires. Les actions disponibles dépendent du type de carte et de l'endroit où vous cliquez.",
"help.6.menuAnnotations": "Ajouter, modifier ou supprimer des annotations.",
"help.6.menuSplits": "Ajouter ou supprimer des changements de phase.",
"help.6.menuExclusions": "Exclure du calcul des limites de contrôle des observations étudiées, ou les réintégrer.",
"help.6.menuCentre": "Afficher, masquer ou rétablir les libellés des lignes centrales.",
"help.6.menuExport": "Copier le graphique, copier son analyse ou exporter une image.",

"help.6.excludeHeading": "Important – exclusion de points de données :",
"help.6.excludeAdvice": "N'excluez une observation du calcul des limites qu'après avoir étudié et compris la raison pour laquelle elle diffère des autres, et déterminé qu'elle ne représente pas la variation habituelle du processus. Les observations exclues restent visibles et peuvent être réintégrées. Pensez à annoter la raison de leur exclusion.",
"help.6.subgroups": "Sur les cartes X̄–S, l'exclusion concerne un sous-groupe entier. Les diagrammes de Pareto ne comportent ni changements de phase des limites de contrôle ni exclusions de points.",
"help.6.splitAdvice": "De même, n'utilisez un changement de phase que lorsque vous avez de bonnes raisons de penser que le processus a changé. Les changements de phase doivent refléter des changements étudiés et compris, et non servir simplement à faire disparaître des signaux gênants.",
"help.6.extraBefore": "Repérez les petits symboles",
"help.6.extraMiddle": "à côté des paramètres dans SimpleSPC. Survolez-les avec la souris ou sélectionnez-les au clavier pour obtenir des conseils supplémentaires. Vous pouvez également interroger",
"help.6.extraAfter": "sur votre carte et son interprétation.",

// AIDE ET CONSEILS : SECTIONS 7 À 11

// 7. Interprétation du résumé
"help.7.title": "7. Interpréter le résumé",
"help.7.intro": "Le résumé situé sous le graphique aide à comprendre les résultats de l'analyse.",
"help.7.commonName": "Variation de cause commune",
"help.7.commonDesc": "signifie que le processus semble stable à son niveau de performance actuel.",
"help.7.specialName": "Variation de cause spéciale",
"help.7.specialDesc": "signifie que le comportement observé suggère qu'un changement a pu se produire.",
"help.7.stableBefore": "Un processus peut être",
"help.7.stableHighlight": "stable mais présenter des performances insuffisantes",
"help.7.stableAfter": "si son niveau de performance reste constamment inférieur au niveau souhaité.",
"help.7.advice": "Recherchez les changements durables plutôt que de vous concentrer sur un seul point favorable ou défavorable. Utilisez votre connaissance du contexte local pour comprendre ce qui se passait à ce moment-là.",

// 8. Enregistrement et réouverture du travail
"help.8.title": "8. Enregistrer et rouvrir votre travail",
"help.8.intro": "Vous pouvez enregistrer votre travail de différentes manières, selon l'utilisation que vous souhaitez en faire par la suite.",
"help.8.saveDesc": "enregistre les données actuelles et les paramètres du graphique dans un fichier de projet .json.",
"help.8.openDesc": "restaure un fichier de projet précédemment enregistré afin que vous puissiez reprendre votre travail ultérieurement.",
"help.8.portableBefore": "crée un fichier autonome",
"help.8.portableAfter": "contenant l'outil SPC, vos données et les paramètres de votre graphique. Ce fichier peut être ouvert directement dans un navigateur et réutilisé sans connexion Internet.",

"help.8.preservedHeading": "Les graphiques enregistrés peuvent notamment conserver :",
"help.8.data": "les données utilisées pour créer le graphique",
"help.8.columns": "les colonnes sélectionnées et le type d'axe X",
"help.8.type": "le type de carte",
"help.8.labels": "les titres, les libellés et les paramètres des axes",
"help.8.rules": "les règles et les objectifs",
"help.8.baseline": "les paramètres de la période de référence",
"help.8.splits": "les changements de phase et les annotations",

"help.8.editableBefore": "Un fichier HTML portable reste modifiable. Après l'avoir ouvert, vous pouvez changer les données ou les paramètres du graphique, puis utiliser à nouveau",
"help.8.editableAfter": "pour créer une copie actualisée.",
"help.8.privacy": "Les fichiers de projet JSON et les fichiers HTML portables contiennent les données utilisées pour créer le graphique. Conservez et partagez ces fichiers conformément aux exigences de gouvernance de l'information, de confidentialité et de cybersécurité de votre organisation.",

// 9. Copie et téléchargement des graphiques
"help.9.title": "9. Copier et télécharger les graphiques",
"help.9.context": "Faites un clic droit sur le graphique pour ouvrir le menu permettant de l'exporter sous forme d'image.",
"help.9.appearance": "Utilisez les options du graphique pour ajuster son apparence avant l'exportation.",
"help.9.clipboard": "Vous pouvez copier les graphiques et/ou leur analyse dans le presse-papiers afin de les coller facilement dans un rapport, une présentation ou un document.",
"help.9.secondaryBefore": "Utilisez l'option",
"help.9.secondaryName": "Afficher/masquer la carte MR/S",
"help.9.secondaryMiddle": "dans l'onglet",
"help.9.secondaryAfter": "pour choisir si la carte secondaire est incluse dans les éléments copiés ou téléchargés.",

// 10. Confidentialité des données
"help.10.title": "10. Confidentialité des données",
"help.10.intro": "L'outil fonctionne entièrement dans votre navigateur :",
"help.10.noServer": "les données ne sont pas envoyées à un serveur",
"help.10.localFiles": "les fichiers importés sont traités localement pendant votre session de navigation",
"help.10.unsavedBefore": "Le travail non enregistré peut être perdu lorsque vous fermez ou actualisez la page. Utilisez",
"help.10.unsavedMiddle": "(JSON) ou",
"help.10.unsavedAfter": "(HTML) pour conserver une copie de votre travail.",
"help.10.caution": "SimpleSPC n'est pas conçu pour traiter des données permettant d'identifier des patients ni d'autres données personnelles sensibles. Utilisez des données anonymisées ou convenablement agrégées et respectez les exigences de gouvernance de l'information, de confidentialité et de cybersécurité de votre organisation. N'oubliez pas que les fichiers JSON et HTML portables contiennent les données utilisées pour créer le graphique.",

// 11. À propos de l'outil
"help.11.title": "11. À propos de l'outil",
"help.11.authorBefore": "Cet outil a été développé par Jem Ramazanoglu –",
"help.11.authorAfter": "comme ressource gratuite pour soutenir l'amélioration de la qualité au sein du NHS.",
"help.11.feedback": "Vos commentaires et suggestions sont les bienvenus.",

//INTERFACE DE L'ASSISTANT SPC
"helper.title": "Assistant SPC",
"helper.close": "Fermer l'assistant SPC",
"helper.closeTitle": "Fermer",
"helper.hideQuestions": "Masquer les questions suggérées",
"helper.showQuestions": "Afficher les questions suggérées",
"helper.generalHeading": "Questions générales sur le SPC",
"helper.chartHeading": "Questions sur mon graphique",
"helper.generalAria": "Questions générales suggérées sur le SPC",
"helper.chartAria": "Questions suggérées sur mon graphique",
"helper.askPlaceholder": "Posez une question…",
"helper.askButton": "Demander",

"helper.introGeneralTitle": "Assistant SPC",
"helper.introGeneral": "Posez une question générale avant de charger des données ou choisissez l'une des questions suggérées ci-dessous.",
"helper.introChartTitle": "Assistant pour le graphique",
"helper.introChart": "Utilisez les questions sur votre graphique pour obtenir une interprétation adaptée à vos données.",
"helper.emptyQuestion": "Saisissez une question ou choisissez une suggestion pour commencer.",

// QUESTIONS GÉNÉRALES SUGGÉRÉES
"helper.q.whatSpc": "Qu'est-ce que le SPC ?",
"helper.q.commonSpecial": "Quelle est la différence entre une variation de cause commune et une variation de cause spéciale ?",
"helper.q.chooseChart": "Comment choisir le bon type de carte ?",
"helper.q.run": "Qu'est-ce qu'une carte de suivi (Run) ?",
"helper.q.xmr": "Qu'est-ce qu'une carte XmR ?",
"helper.q.stable": "Que signifie « stable » ?",
"helper.q.baselineSplit": "Quelle est la différence entre une période de référence et un changement de phase ?",
"helper.q.limits": "Comment fonctionnent les limites de contrôle ?",
"helper.q.signalAction": "Que dois-je faire si je détecte un signal ?",
"helper.q.stablePoor": "Que faire si mon processus est stable, mais que ses performances sont insuffisantes ?",
"helper.q.rulesUnavailable": "Pourquoi certaines règles SPC ne sont-elles pas disponibles ?",
"helper.q.rareRules": "Pourquoi les règles de séquence et de tendance sont-elles désactivées par défaut pour les cartes T et G ?",
"helper.q.zoneRules": "Pourquoi les règles de zone ne sont-elles pas disponibles pour les cartes P, P′, U, U′ et C ?",

// QUESTIONS SUR LE GRAPHIQUE
"helper.q.myChart": "Que m'indique mon graphique ?",
"helper.q.changed": "Quelque chose a-t-il changé ?",
"helper.q.processStable": "Mon processus est-il stable ?",
"helper.q.improvement": "S'agit-il probablement d'une amélioration réelle ?",
"helper.q.target": "Que signifient ces résultats par rapport à notre objectif ?",
"helper.q.next": "Que devrais-je faire ensuite ?",
"helper.q.addSplit": "Dois-je ajouter un changement de phase ici ?",
"helper.q.xmrLimits": "Pourquoi mes limites XmR sont-elles différentes ?",
"helper.q.pPrime": "Pourquoi utilise-t-on une carte P′ ?",
"helper.q.uPrime": "Pourquoi utilise-t-on une carte U′ ?",
"helper.q.sigmaZ": "Que signifie Sigma Z ?",
"helper.q.exclude": "Quand faut-il exclure un point du calcul des limites ?",

// QUESTIONS SUR LE PARETO
"helper.q.pareto80": "Que signifie la ligne des 80 % ?",
"helper.q.paretoCauses": "Les catégories les plus fréquentes sont-elles les causes du problème ?",


  // RÉPONSES GÉNÉRALES DE L'ASSISTANT SPC

  "helper.faq.what-is-spc":
    "La maîtrise statistique des procédés (SPC) permet d'utiliser des données recueillies au fil du temps pour comprendre si un processus se comporte comme d'habitude ou si quelque chose a pu changer.\n\n" +
    "Dans le cadre de l'amélioration de la qualité des soins, elle aide les équipes à distinguer les fluctuations habituelles des signaux pouvant refléter un changement réel dans le service, le parcours de soins, les effectifs, la demande, le codage ou le système de mesure.",

  "helper.faq.common-vs-special":
    "La variation de cause commune correspond aux fluctuations habituelles inhérentes au système actuel. La variation de cause spéciale signifie que le comportement observé suggère qu'un événement inhabituel a pu se produire.\n\n" +
    "En pratique :\n" +
    "• cause commune = le système se comporte comme à son habitude\n" +
    "• cause spéciale = un signal invitant à rechercher ce qui a pu changer",

  "helper.faq.choose-chart":
    "Choisissez la carte selon le type de données dont vous disposez :\n\n" +
    "• Carte de suivi (Run) : données simples suivies au fil du temps, souvent un bon point de départ\n" +
    "• Carte XmR : données continues, avec une mesure par point temporel\n" +
    "• Carte C : nombres d'événements avec des opportunités de survenue relativement constantes\n" +
    "• Carte P : proportions calculées à partir d'un numérateur et d'un dénominateur\n" +
    "• Carte P′ : proportions présentant une variation plus forte ou plus faible que celle attendue pour une carte P standard\n" +
    "• Carte U : taux avec un nombre variable d'opportunités\n" +
    "• Carte U′ : taux présentant une variation plus forte ou plus faible que celle attendue pour une carte U standard\n" +
    "• Carte X̄–S : données continues regroupées en sous-groupes\n" +
    "• Cartes T / G : événements rares\n\n" +
    "• Diagramme de Pareto : classe les catégories selon leur fréquence pour orienter les investigations ; il complète une carte SPC chronologique sans la remplacer\n\n" +
    "Pour les données de proportions, SimpleSPC commence normalement par une carte P et vérifie si une carte P′ correspond mieux à la variation observée. Pour les données de taux avec un nombre variable d'opportunités, l'outil commence normalement par une carte U et vérifie si une carte U′ correspond mieux à la variation observée.",

  "helper.faq.run-chart":
    "Une carte de suivi (Run) présente des données au fil du temps avec une ligne centrale, généralement la médiane. Elle permet de voir simplement si le comportement semble globalement stable ou si un décalage ou une tendance pourrait être présent.\n\n" +
    "C'est souvent un très bon point de départ pour les équipes d'amélioration de la qualité des soins.",

  "helper.faq.xmr-chart":
    "Une carte XmR est utilisée pour des données continues lorsque l'on dispose d'une observation par point temporel.\n\n" +
    "Elle comporte généralement :\n" +
    "• une carte des valeurs individuelles (X), montrant les mesures au fil du temps\n" +
    "• une carte des étendues mobiles (MR), montrant l'écart entre deux points consécutifs\n\n" +
    "Elle aide à distinguer la variation habituelle des éventuels signaux de cause spéciale.",

  "helper.faq.stable-meaning":
    "Un processus stable ne présente que des variations habituelles de cause commune. En termes simples, son comportement est prévisible à son niveau de performance actuel.\n\n" +
    "Stable ne signifie pas nécessairement satisfaisant. Un processus stable peut présenter des performances constamment insuffisantes.",

  "helper.faq.baseline-vs-split":
    "Une période de référence permet de fixer le calcul à partir d'une période initiale. Un changement de phase entraîne un nouveau calcul de la ligne centrale et des limites à partir du point choisi.\n\n" +
    "En termes simples :\n" +
    "• période de référence = comparer les performances ultérieures avec celles de l'ancien système\n" +
    "• changement de phase = considérer la période ultérieure comme un nouveau système ou un nouveau niveau habituel",

  "helper.faq.control-limits":
    "Les limites de contrôle indiquent la plage de valeurs attendue en présence de la variation habituelle de cause commune. Elles ne correspondent pas aux objectifs, aux normes ni aux seuils de réussite ou d'échec.\n\n" +
    "Un point situé au-delà des limites suggère qu'un changement a pu se produire et mérite une investigation. Certaines cartes permettent également d'utiliser des règles sélectionnées pour détecter des comportements particuliers, mais l'outil ne propose ces règles que lorsqu'elles sont raisonnablement justifiables.",

  "helper.faq.what-to-do-signal":
    "Un signal de cause spéciale invite généralement à rechercher ce qui a changé à cette période.\n\n" +
    "Dans un service de santé, cela peut notamment être lié à une modification du parcours de soins, des effectifs, de la demande, du profil des patients, du codage, de la définition des données, à un problème d'équipement ou à une intervention d'amélioration volontaire.\n\n" +
    "L'objectif n'est pas de blâmer une personne ni de réagir excessivement à un point isolé. Il s'agit de comprendre si un élément du système a changé, si ce changement est réel et s'il est susceptible de se maintenir.",

  "helper.faq.stable-but-poor":
    "Si le processus est stable, mais que son niveau de performance est insuffisant, cela signifie généralement que le système produit un résultat prévisible, mais insatisfaisant.\n\n" +
    "Dans une démarche d'amélioration, cela oriente souvent vers une modification ou une refonte du système, plutôt que vers des réactions à chaque bonne ou mauvaise journée.",

  "helper.faq.rule-availability":
    "Certaines règles SPC ne conviennent qu'à des types de cartes précis.\n\n" +
    "Cet outil applique une politique prudente, adaptée au type de carte. Il propose d'abord les règles par défaut les plus sûres et bloque ou masque celles qui pourraient générer des alertes trompeuses.\n\n" +
    "Cette précaution est particulièrement importante dans les services de santé et les services publics, où les faux signaux peuvent entraîner des efforts inutiles, des inquiétudes injustifiées ou une saturation liée aux alertes.",

  "helper.faq.tg-run-trend-off":
    "Les cartes T et G sont des cartes pour événements rares. Elles peuvent naturellement présenter un comportement irrégulier même lorsque le processus n'a pas réellement changé.\n\n" +
    "C'est pourquoi cet outil désactive par défaut les règles de séquence et de tendance pour les cartes T et G. Vous pouvez les activer dans les paramètres avancés, mais leurs résultats doivent être interprétés avec prudence, car ces règles peuvent produire de faux signaux.",

  "helper.faq.puc-zone-rules-blocked":
    "Les règles de zone ne sont pas disponibles pour les cartes P, P′, U, U′ et C, car elles peuvent générer des alertes trompeuses sur ces types de cartes.\n\n" +
    "Ces cartes reposent sur des données d'effectifs, de proportions ou de taux. Leurs limites peuvent varier, et leurs données ne se comportent pas nécessairement comme les données continues pour lesquelles les règles de zone classiques ont été conçues.\n\n" +
    "L'approche par défaut la plus prudente consiste à utiliser les signaux de dépassement des limites et les règles de séquence, plutôt qu'à ajouter des alertes fondées sur les zones.",

  // AUTRES RÉPONSES FIXES DE L'ASSISTANT SPC

  // Calcul des cartes XmR et exclusions manuelles
  "helper.faq.xmr-moving-range-screening":
    "Les limites de contrôle d'une carte XmR estiment la variation habituelle à partir des différences entre les points consécutifs, appelées étendues mobiles.\n\n" +
    "Un changement très important entre deux points peut augmenter l'étendue mobile moyenne et rendre les limites de contrôle anormalement larges. Par défaut, SimpleSPC effectue un filtrage des étendues mobiles exceptionnellement grandes lors de l'estimation de la variation habituelle. Les observations et les étendues mobiles d'origine restent visibles sur le graphique.\n\n" +
    "SimpleSPC utilise la méthode de Nelson : les étendues mobiles supérieures à 3,5 fois l'étendue mobile moyenne initiale ne sont pas utilisées pour estimer la variation. Vous pouvez désactiver ce filtrage dans les options avancées de détection si vous souhaitez utiliser le calcul sans filtrage.",

  "helper.faq.exclude-point-from-limits":
    "N'excluez un point du calcul des limites qu'après l'avoir étudié et si vous avez de bonnes raisons de penser qu'il correspond à une cause spéciale identifiée plutôt qu'à la variation habituelle du processus.\n\n" +
    "Le point n'est pas supprimé. Il reste visible sur le graphique, mais SimpleSPC ne l'utilise pas pour estimer la ligne centrale ou les limites de contrôle. N'excluez pas un point simplement parce que sa valeur est élevée, faible ou située au-delà d'une limite de contrôle. Commencez par rechercher ce qui s'est passé.\n\n" +
    "Si le changement correspond à une nouvelle façon de travailler durable ou à un nouveau niveau de performance du processus, un changement de phase est généralement plus approprié que l'exclusion du point.\n\n" +
    "Pour les cartes XmR, l'exclusion d'une observation entraîne également l'exclusion des étendues mobiles immédiatement avant et après cette observation lors de l'estimation de la variation habituelle. SimpleSPC ne crée pas d'étendue mobile artificielle de part et d'autre du point exclu. Pour les cartes X̄–S, l'exclusion concerne l'ensemble du sous-groupe représenté.\n\n" +
    "Remarque technique : certains ouvrages SPC appellent cette méthode « ghosting » ou « causing out ». Elle permet d'estimer la variation habituelle de cause commune du processus — parfois appelée la voix du processus — sans laisser une cause spéciale étudiée fausser les limites.",

  // Objectifs et périodes de référence
  "helper.faq.target":
    "Un objectif correspond au niveau de performance que vous souhaitez atteindre. En SPC, il doit être interprété en tenant compte de la stabilité du processus.\n\n" +
    "Un processus stable dont les performances sont inférieures à l'objectif nécessite souvent une amélioration ou une refonte du système. Un processus instable qui atteint l'objectif de façon intermittente peut néanmoins manquer de fiabilité.",

  "helper.faq.baseline":
    "Une période de référence est l'ensemble initial de points utilisé pour calculer la ligne centrale et, lorsque cela s'applique, les limites de contrôle.\n\n" +
    "Elle est utile lorsque vous souhaitez comparer les performances ultérieures à une période antérieure représentative du système initial.",

  // Qualité des données et interprétation
  "helper.faq.data-definition-change":
    "Un changement de codage, de critères d'inclusion, de règles de calcul du dénominateur ou de définition des cas peut produire un signal sur le graphique même si le fonctionnement réel du service n'a pas changé.\n\n" +
    "C'est pourquoi les cartes SPC doivent toujours être interprétées en tenant compte des connaissances locales sur la collecte des données et le contexte opérationnel.",

  "helper.faq.annotations":
    "Utilisez les annotations pour signaler des événements pouvant expliquer un changement dans le comportement observé, par exemple un nouveau parcours de soins, une modification des effectifs, un changement de politique ou une nouvelle définition des données.\n\n" +
    "Rédigez des annotations courtes et factuelles. Elles facilitent l'interprétation, mais ne prouvent pas à elles seules un lien de causalité.",

  "helper.faq.improvement-vs-one-good-point":
    "En SPC, une amélioration est plus convaincante lorsque le graphique montre un changement durable dans la direction souhaitée, plutôt qu'un seul bon résultat.\n\n" +
    "Un résultat favorable isolé peut se produire par hasard. Un décalage durable ou un autre signal clair est davantage compatible avec un changement réel.",

  // Explication des lignes centrales
  "helper.faq.runchart-median":
    "Les cartes de suivi (Run) utilisent souvent la médiane, car elle est simple et robuste, notamment lorsque vous commencez une démarche d'amélioration ou disposez de peu de données.\n\n" +
    "Elle fournit une ligne centrale pratique pour examiner les séquences et les tendances.",

  "helper.faq.xmr-mean":
    "Une carte XmR utilise la moyenne et les étendues mobiles pour estimer la variation habituelle et calculer les limites de contrôle.\n\n" +
    "Lorsque les données conviennent à ce type de carte, cela permet de détecter des signaux statistiques plus solides qu'avec une simple carte de suivi.",

  // Changements de phase, objectifs et causalité
  "helper.faq.should-add-split":
    "N'ajoutez généralement un changement de phase que lorsqu'il existe des preuves solides d'un changement réel et durable du processus, et que vous vous attendez à ce que le nouveau comportement se poursuive.\n\n" +
    "Un changement de phase doit faciliter l'interprétation, et non masquer une instabilité ou donner artificiellement une meilleure apparence au graphique.",

  "helper.faq.above-target-solved":
    "Pas nécessairement. Atteindre ou dépasser l'objectif une seule fois ne signifie pas que le processus présente un niveau de performance stable et durable.\n\n" +
    "Il vaut mieux se demander si le processus semble désormais atteindre ou dépasser l'objectif de manière prévisible, et si ce changement est susceptible de durer.",

  "helper.faq.below-target-stable":
    "Cela signifie généralement que le système produit un résultat constant, mais insuffisant.\n\n" +
    "Cette situation oriente souvent vers un travail d'amélioration ou de refonte du système, plutôt que vers des réactions aux points individuels.",

  "helper.faq.caused-improvement":
    "Il faut rester prudent. Un graphique peut montrer que les performances ont changé, et les annotations peuvent aider à établir un lien chronologique avec une intervention, mais le graphique seul ne prouve pas la causalité.\n\n" +
    "Il est généralement préférable de dire que le changement observé est compatible avec l'effet de l'intervention, en s'appuyant sur les connaissances locales.",

  // Profil des patients, dénominateurs et événements rares
  "helper.faq.case-mix":
    "Un changement réel dans le profil ou la complexité des patients peut influencer le graphique, même si le fonctionnement du service est resté identique.\n\n" +
    "Cela ne signifie pas que le graphique est incorrect, mais ce changement doit être pris en compte dans son interprétation.",

  "helper.faq.denominator-changed":
    "Si le nombre d'opportunités varie fortement au fil du temps, cela influence la manière de construire et d'interpréter le graphique.\n\n" +
    "C'est notamment pour cette raison que les cartes P et U sont utiles : elles permettent de tenir compte de totaux ou de nombres d'opportunités variables.",

  "helper.faq.rare-event-improvement":
    "Pour les cartes d'événements rares, une amélioration correspond souvent à un allongement du temps entre les événements indésirables ou à une augmentation du nombre d'opportunités entre ces événements.\n\n" +
    "Comme les données d'événements rares peuvent présenter des variations impressionnantes, il vaut mieux rechercher des changements durables du comportement observé plutôt que de se concentrer sur un seul intervalle exceptionnellement long.",

// ASSISTANT XmR : DESCRIPTIONS DES SIGNAUX
"helper.xmr.signal.above": "un ou plusieurs points au-dessus de la limite supérieure de contrôle",
"helper.xmr.signal.below": "un ou plusieurs points sous la limite inférieure de contrôle",
"helper.xmr.signal.run": "une séquence prolongée d'un même côté de la ligne centrale",
"helper.xmr.signal.trend": "une tendance prolongée",
"helper.xmr.signal.outlier": "une valeur exceptionnellement éloignée",
"helper.xmr.signal.cluster": "un regroupement inhabituel de points éloignés de la ligne centrale",
"helper.xmr.signal.bias": "un comportement durablement éloigné de la ligne centrale",
"helper.xmr.signal.other": "des signaux de cause spéciale",
"helper.xmr.signal.joinTwo": "{first} et {last}",
"helper.xmr.signal.joinMany": "{first} et {last}",

// INTERPRÉTATIONS PAR PÉRIODE
"helper.xmr.period.label": "Période {number}",
"helper.xmr.period.mean": " autour d'une moyenne de {mean}",
"helper.xmr.period.count": "Seulement {count} points ont contribué à l'estimation des limites ; ",
"helper.xmr.period.provisional": "{label} (points {start} à {end}) : les limites XmR sont provisoires{mean}. {count}SimpleSPC exige au moins {minimum} points utilisables avant d'interpréter la stabilité et les signaux de cause spéciale pour cette période",
"helper.xmr.period.stable": "{label} (points {start} à {end}) : le comportement semble stable{mean}",
"helper.xmr.period.signals": "{label} (points {start} à {end}) : signaux possibles de cause spéciale — {signals}{mean}",

// PLUSIEURS PÉRIODES
"helper.xmr.multi.earlier": "Sur l'ensemble de la carte XmR, au moins une période disposant de suffisamment de données présente des signaux de cause spéciale. Cela suggère que le processus a pu changer.\n\n",
"helper.xmr.multi.splits": "Cette carte XmR comporte plusieurs périodes séparées par des changements de phase. Chaque changement de phase entraîne un nouveau calcul de la moyenne et des limites, mais ne prouve pas à lui seul qu'un changement du processus s'est produit.\n\n",
"helper.xmr.multi.count": "Cette période compte actuellement {count} points utilisables et ",
"helper.xmr.multi.provisional": "La période la plus récente (points {start} à {end}) ne contient pas encore assez de données pour conclure à sa stabilité. {count}SimpleSPC exige au moins {minimum} points utilisables avant d'interpréter la stabilité d'une carte XmR.",
"helper.xmr.multi.warning": "\n\nCertains signes précoces méritent une attention particulière, mais cette période reste provisoire. Des observations supplémentaires sont nécessaires pour déterminer si ce comportement persiste.",
"helper.xmr.multi.calm": "\n\nAucun signe précoce évident n'apparaît sur cette courte période. C'est encourageant, mais cela ne suffit pas encore pour conclure à la stabilité.",
"helper.xmr.multi.stable": "Au cours de la période la plus récente (points {start} à {end}), aucun signal clair de cause spéciale n'est détecté. Le nombre de données étant suffisant, ce comportement est compatible avec la variation habituelle au niveau de performance actuel.",
"helper.xmr.multi.signals": "Au cours de la période la plus récente (points {start} à {end}), des signes possibles de cause spéciale persistent : {signals}.",

// PÉRIODE UNIQUE
"helper.xmr.single.noChart": "Générez d'abord une carte XmR, puis reposez votre question.",
"helper.xmr.single.count": "Actuellement, {count} points utilisables contribuent au calcul des limites et ",
"helper.xmr.single.provisional": "Les données sont encore insuffisantes pour conclure à la stabilité du processus. {count}SimpleSPC exige au moins {minimum} points utilisables pour une interprétation formelle de la carte XmR.",
"helper.xmr.single.warning": "\n\nCertains éléments méritent d'être surveillés, mais davantage d'observations sont nécessaires pour déterminer s'ils constituent un signal durable.",
"helper.xmr.single.calm": "\n\nLe comportement observé semble actuellement relativement régulier, sans signe précoce évident. Toutefois, les conclusions restent provisoires jusqu'à ce que suffisamment d'observations soient disponibles.",
"helper.xmr.single.stable": "La carte XmR actuelle contient suffisamment de données pour être interprétée et ne présente aucun signal clair de cause spéciale. Ce comportement est compatible avec la variation habituelle, même si une performance stable n'est pas nécessairement une performance satisfaisante.",
"helper.xmr.single.signals": "La carte XmR actuelle suggère que le processus a pu changer. Le comportement observé ne s'explique pas entièrement par la variation habituelle. Il convient donc d'étudier ce qui s'est passé dans le processus ou dans son système de mesure.",

  // ÉVALUATION DE LA STABILITÉ

  // Noms des cartes dans les réponses dynamiques
  "helper.stability.chart.c": "carte C",
  "helper.stability.chart.p": "carte P",
  "helper.stability.chart.pprime": "carte P′",
  "helper.stability.chart.u": "carte U",
  "helper.stability.chart.uprime": "carte U′",
  "helper.stability.chart.t": "carte T",
  "helper.stability.chart.g": "carte G",

  "helper.stability.signalFallback":
    "un ou plusieurs signaux SPC",

  // Pareto
  "helper.stability.pareto":
    "Un diagramme de Pareto ne permet pas d'évaluer la stabilité d'un processus, car les catégories ne sont pas représentées dans l'ordre chronologique.",

  // XmR : identification de la période
  "helper.stability.xmr.latest":
    "Pour la période la plus récente (période {period} sur {total}, points {start} à {end})",

  "helper.stability.xmr.single":
    "Pour le graphique actuel (points {start} à {end})",

  // XmR : limites provisoires
  "helper.stability.xmr.count":
    "Cette période compte actuellement {count} points et ",

  "helper.stability.xmr.provisional":
    "{scope} : les données sont encore insuffisantes pour conclure à la stabilité du processus. {count}SimpleSPC exige au moins {minimum} points avant de procéder à une interprétation formelle de la stabilité XmR.",

  "helper.stability.xmr.warning":
    "\n\nCertains éléments de cette courte période méritent déjà une attention particulière. Il est donc trop tôt pour qualifier le processus de stable. Des données supplémentaires permettront de déterminer s'il s'agit de fluctuations temporaires ou d'un comportement durable.",

  "helper.stability.xmr.calm":
    "\n\nJusqu'à présent, le comportement semble relativement régulier, sans signe d'alerte évident sur cette courte période. C'est encourageant, mais les conclusions restent provisoires et ne prouvent pas encore la stabilité du processus.",

  // XmR : périodes pouvant être interprétées
  "helper.stability.xmr.stable":
    "{scope} : aucun signal clair de cause spéciale n'est détecté selon les règles actuellement sélectionnées. Cela suggère que le processus présente un comportement relativement régulier autour de son niveau de performance actuel.",

  "helper.stability.xmr.signals":
    "{scope} : cette carte XmR présente des signes possibles de variation de cause spéciale, notamment {signals}.",

  // Cartes X-bar et S
  "helper.stability.xbars.bothStable":
    "Au cours de la période X̄–S la plus récente, les moyennes des sous-groupes et la variation à l'intérieur des sous-groupes semblent toutes deux stables. Aucun signal clair de cause spéciale n'est détecté dans l'une ou l'autre des cartes.",

  "helper.stability.xbars.xChanged":
    "Au cours de la période X̄–S la plus récente, les moyennes des sous-groupes présentent des signes possibles de changement, tandis que la variation à l'intérieur des sous-groupes semble stable.",

  "helper.stability.xbars.sChanged":
    "Au cours de la période X̄–S la plus récente, les moyennes des sous-groupes semblent stables, tandis que la variation à l'intérieur des sous-groupes présente des signes possibles de changement.",

  "helper.stability.xbars.bothChanged":
    "Au cours de la période X̄–S la plus récente, les moyennes des sous-groupes et la variation à l'intérieur des sous-groupes présentent toutes deux des signes possibles de changement.",

  // Cartes aux attributs
  "helper.stability.attribute.stable":
    "Pour la période la plus récente de la {chart}, aucun signal clair de cause spéciale n'est détecté selon les règles actuellement sélectionnées.",

  "helper.stability.attribute.signals":
    "Pour la période la plus récente de la {chart}, des signes possibles de variation de cause spéciale sont détectés, notamment {signals}.",

  // Cartes d'événements rares
  "helper.stability.rare.stable":
    "Pour la période la plus récente de la {chart}, aucun signal clair de cause spéciale n'est détecté selon les règles actuellement sélectionnées pour les événements rares.",

  "helper.stability.rare.signals":
    "Pour la période la plus récente de la {chart}, des signes possibles de variation de cause spéciale sont détectés, notamment {signals}.",

  "helper.stability.noChart":
    "Générez d'abord la carte actuellement sélectionnée pour obtenir une interprétation adaptée à vos données.",

  // Cartes de suivi (Run)
  "helper.stability.run.noChart":
    "Générez d'abord une carte de suivi (Run) pour obtenir une interprétation.",

  "helper.stability.run.splitIntro":
    "Cette carte de suivi (Run) a été divisée en {count} périodes par des changements de phase. Chaque période possède sa propre médiane et SimpleSPC recherche séparément les décalages et les tendances dans chaque période.\n\n",

  "helper.stability.run.splitStable":
    "Aucun décalage ni aucune tendance n'ont été détectés à l'intérieur des différentes périodes selon les règles actuellement sélectionnées. Cela ne signifie toutefois pas que le processus est resté inchangé dans son ensemble.\n\n",

  "helper.stability.run.splitSignals":
    "Une ou plusieurs périodes présentent un signal de décalage ou de tendance qui pourrait nécessiter une investigation.\n\n",

  "helper.stability.run.splitAdvice":
    "Comparez les médianes indiquées pour les différentes périodes. Une différence de niveau peut révéler un changement important de performance, même si aucun autre signal n'apparaît à l'intérieur des périodes.\n\nUn changement de phase correspond à une séparation choisie pour l'analyse. Il ne prouve pas à lui seul qu'un changement statistiquement significatif s'est produit. Examinez pourquoi ce changement de phase a été introduit, ce qui s'est passé dans le processus à ce moment-là et si le nouveau niveau de performance se maintient.",

  "helper.stability.run.stable":
    "Sur la carte de suivi (Run) actuelle, aucun décalage ni aucune tendance n'ont été détectés selon les règles sélectionnées.",

  "helper.stability.run.signals":
    "La carte de suivi (Run) actuelle présente un ou plusieurs signaux de décalage ou de tendance, suggérant que le comportement du processus a pu changer. Recherchez ce qui s'est passé dans le processus à cette période.",

  // SIGNAUX DYNAMIQUES DE L'ASSISTANT
  "helper.signal.shiftMean": "un décalage prolongé (de nombreux points du même côté de la moyenne)",
  "helper.signal.shiftDetailed": "Décalage : {count}+ points consécutifs {direction} la ligne centrale (de {start} à {end})",
  "helper.signal.trendDetailed": "Tendance : {count}+ points {direction} de façon continue (de {start} à {end})",
  "helper.signal.sideAbove": "au-dessus de",
  "helper.signal.sideBelow": "au-dessous de",
  "helper.signal.directionIncreasing": "en augmentation",
  "helper.signal.directionDecreasing": "en diminution",

  // OBJECTIFS ET AMÉLIORATION : RÉPONSES DYNAMIQUES

  // Noms complémentaires des cartes
  "helper.stability.chart.run": "carte de suivi (Run)",
  "helper.stability.chart.xmr": "carte XmR",
  "helper.stability.chart.xbars": "carte X̄–S",
  "helper.stability.chart.pareto": "diagramme de Pareto",

  // XmR : absence d'objectif
  "helper.target.xmr.none":
    "Aucun objectif n'est actuellement défini pour cette carte XmR. Si vous en ajoutez un, SimpleSPC pourra vous aider à interpréter les performances par rapport à cet objectif, tout en tenant compte de la stabilité du processus.",

  // XmR : moyenne et objectif
  "helper.target.xmr.meanFavourable":
    "La moyenne de la période actuelle ({mean}) se situe du côté favorable de l'objectif ({target}).",
  "helper.target.xmr.meanUnfavourable":
    "La moyenne de la période actuelle ({mean}) ne se situe pas encore du côté favorable de l'objectif ({target}).",

  // XmR : limites provisoires
  "helper.target.xmr.provisional":
    "Les limites de la période XmR la plus récente sont encore provisoires. Il est donc trop tôt pour déterminer avec confiance si le processus peut atteindre régulièrement l'objectif.",
  "helper.target.xmr.count":
    "Cette période compte actuellement {count} points et ",
  "helper.target.xmr.minimum":
    "{count}SimpleSPC exige au moins {minimum} points avant de procéder à une interprétation formelle de la stabilité ou de la capabilité.",
  "helper.target.xmr.earlyFavourable":
    "Les premiers résultats sont encourageants, mais ils ne démontrent pas encore que ce niveau de performance est prévisible ou durable.",
  "helper.target.xmr.earlyUnfavourable":
    "Le comportement initial peut encore évoluer. Des données supplémentaires permettront de déterminer si le niveau du processus a réellement changé et si d'autres améliorations sont nécessaires pour atteindre durablement l'objectif.",

  // XmR : stabilité et causes spéciales
  "helper.target.xmr.stable":
    "La période XmR la plus récente dispose de suffisamment de données pour être interprétée et semble actuellement stable.",
  "helper.target.xmr.capability":
    "Sur la base du processus stable actuel, SimpleSPC estime qu'environ {percentage} % des résultats futurs atteindraient l'objectif, à condition que le processus continue à se comporter de la même manière.",
  "helper.target.xmr.stableNoCapability":
    "Comme le processus semble stable, la comparaison avec l'objectif est plus pertinente. La question essentielle est de savoir si le système actuel peut atteindre de manière prévisible un niveau de performance satisfaisant.",
  "helper.target.xmr.unstable":
    "La période XmR la plus récente dispose de suffisamment de données pour être interprétée, mais elle présente une variation de cause spéciale. Les résultats par rapport à l'objectif doivent donc être interprétés avec prudence, car le processus ne se comporte pas encore de façon prévisible.",
  "helper.target.xmr.unstableAction":
    "La priorité est de comprendre ce qui change dans le système, la demande, les effectifs, le parcours de soins, le codage ou les mesures avant de s'appuyer sur une estimation de capabilité.",

  // Autres cartes : interprétation de l'objectif
  "helper.target.general.none":
    "Aucun objectif n'est actuellement défini pour le graphique sélectionné ({chart}). Si vous en ajoutez un, SimpleSPC pourra vous aider à l'interpréter en tenant compte de la variation observée.",
  "helper.target.general.higher":
    "les valeurs plus élevées sont considérées comme préférables",
  "helper.target.general.lower":
    "les valeurs plus faibles sont considérées comme préférables",
  "helper.target.general.intro":
    "L'objectif est fixé à {target} et {direction}.",
  "helper.target.general.centreFavourable":
    "La ligne centrale actuelle ({centre}) se situe du côté favorable de l'objectif.",
  "helper.target.general.centreUnfavourable":
    "La ligne centrale actuelle ({centre}) ne se situe pas encore du côté favorable de l'objectif.",
  "helper.target.general.stable":
    "La période la plus récente ne présente aucun signal clair de cause spéciale. La comparaison avec l'objectif est donc plus pertinente. Rappelez-vous qu'un processus stable n'offre pas nécessairement des performances satisfaisantes.",
  "helper.target.general.unstable":
    "La période la plus récente présente des signes possibles de variation de cause spéciale. Les performances par rapport à l'objectif doivent donc être interprétées avec prudence. Il faut d'abord comprendre ce qui a pu changer avant de supposer que le niveau actuel se maintiendra.",
  "helper.target.general.unknown":
    "Interprétez l'objectif en tenant compte du comportement global du processus, plutôt que de juger chaque point isolément.",

  // Amélioration : XmR
  "helper.improvement.xmr.provisional":
    "Les données de la période XmR la plus récente sont encore insuffisantes pour conclure à la stabilité ou à une amélioration. Le comportement initial peut être encourageant, mais davantage de points sont nécessaires avant de parler d'un changement durable.",
  "helper.improvement.xmr.stable":
    "La période XmR actuelle ne présente aucun signal clair de cause spéciale. Le graphique seul ne fournit donc pas de preuve solide d'une nouvelle amélioration. Une amélioration est plus convaincante lorsqu'un changement durable se produit dans la direction souhaitée et que le nouveau comportement se maintient.",
  "helper.improvement.xmr.unstable":
    "La carte XmR présente des signes suggérant que le processus a pu changer. Cela ne signifie pas automatiquement qu'il s'agit d'une amélioration. Vérifiez si le changement va dans la direction souhaitée, s'il est durable et si le contexte local fournit une explication plausible. Le graphique peut montrer un changement, mais ne prouve pas qu'une intervention en est la cause.",

  // Amélioration : autres cartes
  
  "helper.improvement.general.stable":
    "Sur le graphique actuel ({chart}), aucun signal clair de cause spéciale n'est détecté. Le graphique seul ne fournit donc pas de preuve solide d'une nouvelle amélioration. Une amélioration est plus convaincante lorsqu'un changement durable se produit dans la direction souhaitée, plutôt qu'un seul point favorable.",
  "helper.improvement.general.unstable":
    "Le graphique actuel ({chart}) présente des signes suggérant que le processus a pu changer. Il peut s'agir d'une amélioration, d'une détérioration ou d'un autre changement du système. Pour parler d'amélioration, vérifiez que le changement va dans la direction souhaitée, qu'il est durable et qu'il est cohérent avec le contexte local. Le graphique seul ne prouve pas la causalité.",
  "helper.improvement.general.noChart":
    "Générez d'abord la carte actuellement sélectionnée pour que je puisse vous aider à déterminer si le comportement observé correspond à une amélioration réelle.",

  // LIMITES XmR ET FILTRAGE DE NELSON

  "helper.xmrLimits.noChart":
    "Les limites d'une carte XmR sont estimées à partir des étendues mobiles entre points consécutifs. Générez une carte XmR pour obtenir une explication du calcul appliqué à vos données actuelles.",

  "helper.xmrLimits.intro":
    "Les limites de contrôle XmR dépendent de l'estimation de la variation habituelle. SimpleSPC estime cette variation à partir des étendues mobiles entre observations consécutives.",

  "helper.xmrLimits.screeningOn":
    "\n\nPour votre graphique actuel, le filtrage des étendues mobiles est ACTIVÉ.",

  "helper.xmrLimits.screenedOne":
    " SimpleSPC a repéré {count} variation exceptionnellement importante entre deux points consécutifs et ne l'a pas utilisée pour estimer la variation habituelle. Les points et les étendues mobiles d'origine restent visibles sur le graphique.",

  "helper.xmrLimits.screenedMany":
    " SimpleSPC a repéré {count} variations exceptionnellement importantes entre des points consécutifs et ne les a pas utilisées pour estimer la variation habituelle. Les points et les étendues mobiles d'origine restent visibles sur le graphique.",

  "helper.xmrLimits.whyScreen":
    "\n\nCe filtrage évite qu'un saut exceptionnellement important n'augmente artificiellement l'estimation de la variation et, par conséquent, n'élargisse excessivement les limites de contrôle XmR.",

  "helper.xmrLimits.noneScreened":
    " Aucune étendue mobile de la période actuelle ne dépasse le seuil de filtrage. Le filtrage n'a donc pas modifié l'estimation de la variation pour cette période.",

  "helper.xmrLimits.statistics":
    "\n\nPour la période actuelle :\n" +
    "• moyenne initiale des étendues mobiles = {initial}\n" +
    "• seuil de filtrage de Nelson = {threshold}\n" +
    "• moyenne des étendues mobiles utilisée pour calculer les limites = {average}",

  "helper.xmrLimits.screeningOff":
    "\n\nPour votre graphique actuel, le filtrage des étendues mobiles est DÉSACTIVÉ. Toutes les étendues mobiles de la période de référence sont donc utilisées pour estimer la variation habituelle. Si un écart entre deux points consécutifs est exceptionnellement important, il peut augmenter l'étendue mobile moyenne et produire des limites de contrôle plus larges.",

  "helper.xmrLimits.offAverage":
    "\n\nL'étendue mobile moyenne actuellement utilisée pour calculer les limites est de {average}.",

  "helper.xmrLimits.unknown":
    "\n\nIl n'est pas possible de déterminer si le filtrage des étendues mobiles est activé à partir des informations disponibles pour le graphique actuel.",

  "helper.xmrLimits.splits":
    "\n\nCe graphique comporte également un ou plusieurs changements de phase. Chaque période est analysée séparément, avec sa propre moyenne et ses propres limites. L'étendue mobile située de part et d'autre d'un changement de phase n'est pas utilisée. Cela peut également expliquer pourquoi les limites diffèrent de celles obtenues sans changement de phase.",

  "helper.xmrLimits.provisionalIntro":
    "\n\nLes limites de la période la plus récente sont encore provisoires. ",

  "helper.xmrLimits.provisionalCount":
    "Seulement {count} points sont actuellement utilisés pour estimer ses limites. ",

  "helper.xmrLimits.provisionalMinimum":
    "SimpleSPC attend qu'au moins {minimum} points soient disponibles avant de procéder à une interprétation formelle de la stabilité.",

//  CARTES P′, U′ ET FACTEUR SIGMA Z

// Explications de la carte P′
"helper.laney.pprime.noChart":
  "La carte P′ est une version ajustée de la carte P selon la méthode de Laney. Elle est utilisée pour les données de proportions lorsque la variation observée diffère sensiblement de celle attendue avec une carte P standard.",

"helper.laney.pprime.intro":
  "Cette carte P′ utilise les mêmes numérateurs, dénominateurs, proportions observées et ligne centrale qu'une carte P standard. Elle ajuste uniquement la largeur des limites de contrôle pour mieux refléter la variation effectivement observée dans les données.",

"helper.laney.pprime.more":
  "\n\nVos données varient davantage que ne le prévoit une carte P standard. La carte P′ utilise donc des limites de contrôle plus larges. Cela évite de considérer à tort une variation habituelle comme un changement possible du processus.",

"helper.laney.pprime.less":
  "\n\nVos données varient moins que ne le prévoit une carte P standard. La carte P′ utilise donc des limites de contrôle plus étroites. Cela peut faciliter la détection de variations réellement inhabituelles.",

"helper.laney.pprime.equal":
  "\n\nLa variation observée est très proche de celle attendue avec une carte P standard. L'ajustement de la carte P′ est donc très faible.",

"helper.laney.pprime.factor":
  "\n\nLe facteur d'ajustement Sigma Z de cette carte est de {sigmaZ}. Une valeur supérieure à 1 élargit les limites de la carte P standard ; une valeur inférieure à 1 les rétrécit.",

// Explications de la carte U′
"helper.laney.uprime.noChart":
  "La carte U′ est une version ajustée de la carte U selon la méthode de Laney. Elle est utilisée pour les données de taux lorsque la variation observée diffère sensiblement de celle attendue avec une carte U standard.",

"helper.laney.uprime.intro":
  "Cette carte U′ utilise les mêmes nombres d'événements, opportunités, taux observés et ligne centrale qu'une carte U standard. Elle ajuste uniquement la largeur des limites de contrôle pour mieux refléter la variation effectivement observée dans les données.",

"helper.laney.uprime.more":
  "\n\nVos données varient davantage que ne le prévoit une carte U standard. La carte U′ utilise donc des limites de contrôle plus larges. Cela évite de considérer à tort une variation habituelle comme un changement possible du processus.",

"helper.laney.uprime.less":
  "\n\nVos données varient moins que ne le prévoit une carte U standard. La carte U′ utilise donc des limites de contrôle plus étroites. Cela peut faciliter la détection de variations réellement inhabituelles.",

"helper.laney.uprime.equal":
  "\n\nLa variation observée est très proche de celle attendue avec une carte U standard. L'ajustement de la carte U′ est donc très faible.",

"helper.laney.uprime.factor":
  "\n\nLe facteur d'ajustement Sigma Z de cette carte est de {sigmaZ}. Une valeur supérieure à 1 élargit les limites de la carte U standard ; une valeur inférieure à 1 les rétrécit.",

// Noms des cartes pour les explications techniques
"helper.laney.chart.p": "carte P",
"helper.laney.chart.u": "carte U",
"helper.laney.chart.pprime": "carte P′",
"helper.laney.chart.uprime": "carte U′",

// Explications du facteur Sigma Z
"helper.laney.sigma.general":
  "Sigma Z est utilisé sur les cartes P′ et U′ de Laney. Ce facteur d'ajustement indique si les données présentent une variation plus forte ou plus faible que celle normalement attendue avec une carte P ou U standard.",

"helper.laney.sigma.intro":
  "Sigma Z est le facteur d'ajustement utilisé par une carte de Laney. Il indique si vos données varient davantage ou moins que ne le prévoit une {standard} standard.",

"helper.laney.sigma.guidance":
  "\n\nUne valeur de Sigma Z proche de 1 signifie que la variation observée correspond approximativement à celle attendue avec la carte standard. Une valeur supérieure à 1 indique une variation plus importante, ce qui élargit les limites de contrôle. Une valeur inférieure à 1 indique une variation plus faible, ce qui rétrécit les limites de contrôle.\n\n",

"helper.laney.sigma.unchanged":
  "Sigma Z ne modifie ni les observations d'origine ni la ligne centrale. Il ajuste uniquement la largeur des limites de contrôle sur la {prime}.",

"helper.laney.sigma.value":
  "\n\nPour votre graphique actuel, Sigma Z = {sigmaZ}.",

"helper.laney.sigma.more":
  " Cela signifie que les données varient davantage que ne le prévoit une {standard} standard. La {prime} utilise donc des limites plus larges.",

"helper.laney.sigma.less":
  " Cela signifie que les données varient moins que ne le prévoit une {standard} standard. La {prime} utilise donc des limites plus étroites.",

"helper.laney.sigma.equal":
  " Cette valeur est exactement égale à 1. L'ajustement de Laney ne modifie donc pas la largeur des limites de la {standard} standard.",

"helper.laney.sigma.technical":
  "\n\nDétail technique : SimpleSPC tient d'abord compte de la variation attendue pour chaque point selon la carte standard. Il examine ensuite les changements entre valeurs standardisées consécutives. Sigma Z résume cette variation supplémentaire ou réduite en un seul facteur d'ajustement.",

  //  CONSEILS SUR LES COLONNES ET SUITE À DONNER

  "chart.columnWarning.c":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte C, la colonne des valeurs doit contenir un nombre d'événements ou de défauts par période. N'utilisez pas de colonne de dénominateur.",

  "chart.columnWarning.p":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte P, la colonne des valeurs doit contenir le numérateur, par exemple le nombre de cas ou d'événements, et la troisième colonne doit contenir le dénominateur, par exemple le nombre total de patients ou de séjours.",

  "chart.columnWarning.pprime":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte P′, la colonne des valeurs doit contenir le numérateur, par exemple le nombre de cas ou d'événements, et la troisième colonne doit contenir le dénominateur, par exemple le nombre total de patients ou de séjours. La carte P′ utilise les mêmes données qu'une carte P, mais ajuste les limites de contrôle selon la méthode de Laney.",

  "chart.columnWarning.u":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte U, la colonne des valeurs doit contenir le nombre de défauts ou d'événements, et la troisième colonne doit contenir le nombre d'opportunités ou le dénominateur, par exemple les journées d'hospitalisation, les consultations ou les interventions.",

  "chart.columnWarning.uprime":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte U′, la colonne des valeurs doit contenir le nombre de défauts ou d'événements, et la troisième colonne doit contenir le nombre d'opportunités ou le dénominateur, par exemple les journées d'hospitalisation, les consultations, les interventions ou les inspections. La carte U′ utilise les mêmes données qu'une carte U, mais ajuste les limites de contrôle selon la méthode de Laney.",

  "chart.columnWarning.xbars":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte X̄–S, la colonne des valeurs doit contenir les mesures individuelles, et la troisième colonne doit identifier le sous-groupe, par exemple le jour, la semaine, le lot ou l'échantillon.",

  "chart.columnWarning.tGaps":
    "<strong>Carte T : utilisation d'intervalles numériques.</strong><br>Sélectionnez la colonne contenant le temps entre deux événements consécutifs (par exemple, en jours) comme <strong>colonne des valeurs / axe Y</strong>. La colonne de l'axe X fournit les libellés des observations.<br><br><strong>Vous préférez utiliser les dates des événements ?</strong> Cliquez sur <strong>Paramètres du graphique</strong>, puis sélectionnez <strong>Je dispose des dates des événements</strong>.",

  "chart.columnWarning.tDates":
    "<strong>Carte T : utilisation des dates des événements.</strong><br>Sélectionnez les dates ou heures des événements dans la <strong>colonne des dates / axe X</strong>. SimpleSPC calcule automatiquement les intervalles ; la colonne des valeurs / axe Y est donc désactivée.<br><br><strong>Vous disposez déjà des intervalles calculés ?</strong> Cliquez sur <strong>Paramètres du graphique</strong>, puis sélectionnez <strong>Je dispose déjà des intervalles</strong>.",

  "chart.columnWarning.g":
    "<strong>Vérifiez les colonnes sélectionnées.</strong> Pour une carte G, la colonne des valeurs doit généralement représenter le nombre d'opportunités, de cas ou d'observations entre des événements rares. Vérifiez qu'il ne s'agit pas d'un taux ou d'un pourcentage.",

  "chart.columnWarning.pareto":
    "<strong>Choisissez comment regrouper les occurrences.</strong> Sélectionnez la catégorie à comparer. Si chaque ligne correspond à une occurrence, choisissez <strong>(compter chaque ligne comme 1)</strong> pour la colonne des fréquences. Si les données sont déjà regroupées, sélectionnez plutôt la colonne contenant les effectifs numériques. Les catégories de Pareto aident à orienter les investigations, mais ne doivent pas être considérées automatiquement comme les causes du problème.",

  "helper.decision.general":
    "Dans une démarche d'amélioration de la qualité des soins, les prochaines étapes sont généralement les suivantes :\n" +
    "• si le processus semble stable, mais que ses performances sont insuffisantes, envisagez d'améliorer ou de repenser le système\n" +
    "• si le graphique présente un signal, recherchez ce qui a changé dans le service, le parcours de soins, la demande, les effectifs, le codage ou le système de mesure\n" +
    "• évitez de réagir à chaque point individuel comme s'il prouvait à lui seul une réussite ou un échec",

  "helper.decision.paretoNext":
    "\n\nUne prochaine étape pertinente consiste à étudier plus en détail les catégories les plus fréquentes, en s'appuyant sur les connaissances locales et, si nécessaire, sur une analyse plus détaillée des sous-catégories. Ne supposez pas que la catégorie elle-même représente la cause sous-jacente.",

  // INTERPRÉTATIONS DU DIAGRAMME DE PARETO

  "helper.pareto.noChart":
    "Un diagramme de Pareto classe les catégories de la plus fréquente à la moins fréquente. Il permet de déterminer où des investigations ou des efforts d'amélioration pourraient être les plus utiles.",

  "helper.pareto.firstCategory":
    "la première catégorie",

  "helper.pareto.summary":
    "Ce diagramme de Pareto contient {total} occurrences réparties dans {categories} catégories. Les catégories sont classées de la plus fréquente à la moins fréquente.",

  "helper.pareto.topOne":
    "La catégorie la plus fréquente est « {category} », avec {count} occurrence{share}.",

  "helper.pareto.topMany":
    "La catégorie la plus fréquente est « {category} », avec {count} occurrences{share}.",

  "helper.pareto.topShare":
    " ({percentage} % du total)",

  "helper.pareto.cutoffOne":
    "La première catégorie représente environ {percentage} % des occurrences. Cela peut aider à déterminer les investigations prioritaires, mais la ligne des 80 % est un outil de priorisation et non une limite statistique.",

  "helper.pareto.cutoffMany":
    "Les {count} premières catégories représentent environ {percentage} % des occurrences. Cela peut aider à déterminer les investigations prioritaires, mais la ligne des 80 % est un outil de priorisation et non une limite statistique.",

  "helper.pareto.smallSample":
    "Ce diagramme contient moins de 30 occurrences. Le classement des catégories pourrait donc changer assez facilement à mesure que de nouvelles données seront recueillies.",

  "helper.pareto.causeCaution":
    "Ne supposez pas que les catégories les plus fréquentes correspondent nécessairement aux causes du problème. Elles représentent des façons de regrouper les données et doivent généralement conduire à des investigations supplémentaires.",

  "helper.pareto.notTimeOrdered":
    "Un diagramme de Pareto ne présente pas les données dans l'ordre chronologique. Il ne permet donc pas de déterminer si un processus est stable ni si un changement est statistiquement significatif. Utilisez également une carte de suivi (Run) ou une carte de Shewhart lorsque cette question est importante.",

  // LIGNE DE RÉFÉRENCE À 80 %

  "helper.pareto.eighty.explain":
    "La ligne pointillée des 80 % sert à orienter les investigations prioritaires. Ce n'est pas une limite de contrôle statistique. La courbe cumulative rouge montre comment le pourcentage des occurrences augmente lorsque les catégories sont ajoutées de gauche à droite.",

  "helper.pareto.eighty.countOne":
    "Sur votre diagramme, la première catégorie représente {percentage} % de toutes les occurrences.",

  "helper.pareto.eighty.countMany":
    "Sur votre diagramme, les {count} premières catégories représentent {percentage} % de toutes les occurrences.",

  "helper.pareto.eighty.notRule":
    "La ligne des 80 % constitue un repère utile. Elle ne signifie pas que 80 % des problèmes doivent nécessairement provenir de 20 % des catégories.",

  // LES CATÉGORIES LES PLUS FRÉQUENTES SONT-ELLES LES CAUSES ?

  "helper.pareto.causes":
    "Pas nécessairement. Un diagramme de Pareto montre les catégories qui regroupent le plus grand nombre d'occurrences enregistrées, mais il ne permet pas d'établir ce qui les a causées.\n\n" +
    "Les catégories les plus fréquentes constituent souvent de bons points de départ pour les investigations. Examinez plus précisément ce qui se passe dans ces catégories, tenez compte de la manière dont les données ont été recueillies et utilisez les connaissances locales pour identifier les causes possibles.\n\n" +
    "Une catégorie fréquente n'est pas automatiquement une cause profonde. Une catégorie moins fréquente peut également mériter une attention prioritaire si ses conséquences sont particulièrement graves.",

  // UN CHANGEMENT S'EST-IL PRODUIT ?

  "helper.pareto.changeIntro":
    "Un diagramme de Pareto ne permet pas de déterminer si quelque chose a changé au fil du temps, car il supprime volontairement l'ordre chronologique pour classer les catégories selon leur fréquence.\n\n",

  // RÉSUMÉ STATISTIQUE PARETO

  "summary.pareto.largestLabel": "Catégorie la plus fréquente :",
  "summary.pareto.topThreeLabel": "Trois catégories principales :",
  "summary.pareto.eightyHeading": "Repère des 80 % :",
  "summary.pareto.cautionLabel": "Attention :",
  "summary.pareto.dataCheckLabel": "Vérification des données :",
  "summary.pareto.interpretationLabel": "Interprétation :",
  "summary.pareto.spcNoteLabel": "Remarque SPC :",

  "summary.pareto.occurrenceOne": "occurrence",
  "summary.pareto.occurrenceMany": "occurrences",
  "summary.pareto.ofTotal": "du total",

  "summary.pareto.cutoffOneStart": "la première catégorie",
  "summary.pareto.cutoffOneEnd": "représente",
  "summary.pareto.cutoffManyStart": "les",
  "summary.pareto.cutoffManyEnd": "premières catégories représentent",
  "summary.pareto.ofOccurrences": "des occurrences.",

  "summary.pareto.eightyCaution":
    "La ligne pointillée des 80 % sert à définir des priorités d'investigation ; elle ne constitue pas une limite de contrôle statistique.",

  "summary.pareto.smallSample":
    "ce diagramme de Pareto contient moins de 30 occurrences. Avec un faible nombre d'observations, le classement apparent des catégories peut être instable. Il doit donc être interprété avec prudence.",

  "summary.pareto.blankOne":
    "ligne a été ignorée parce que la catégorie était vide.",

  "summary.pareto.blankMany":
    "lignes ont été ignorées parce que la catégorie était vide.",

  "summary.pareto.interpretation":
    "les catégories situées à gauche sont les plus fréquentes et peuvent constituer des priorités pour les investigations ou les actions d'amélioration. Une catégorie ne représente pas automatiquement une cause du problème.",

  "summary.pareto.spcNote":
    "un diagramme de Pareto ne présente pas les données dans l'ordre chronologique et ne permet pas de déterminer si la variation relève de causes communes ou spéciales. Lorsque cela est pertinent, interprétez-le parallèlement à une carte de suivi (Run) ou à une carte de Shewhart adaptée.",

  //  CHANGEMENTS DU PROCESSUS ET RÈGLES SPC

  // X-bar et S : quatre situations possibles
  "helper.change.xbars.bothStable":
    "Aucun élément clair ne suggère que le processus a changé au cours de la période X̄–S la plus récente. Ni les moyennes des sous-groupes ni la variation à l'intérieur des sous-groupes ne présentent de signal de cause spéciale.",

  "helper.change.xbars.xOnly":
    "Certains signes suggèrent que le niveau du processus a pu changer : la carte X̄ présente un signal, tandis que la carte S reste stable.",

  "helper.change.xbars.sOnly":
    "Certains signes suggèrent que la variation du processus a pu changer : la carte S présente un signal, tandis que la carte X̄ reste stable.",

  "helper.change.xbars.bothSignals":
    "Certains signes suggèrent que le processus a pu changer : les cartes X̄ et S présentent toutes deux des signaux de cause spéciale.",

  // Évaluation des changements : cartes aux attributs
  "helper.change.attribute.stable":
    "Pour la période la plus récente du graphique ({chart}), aucun élément clair ne suggère un changement du processus. Aucun signal de cause spéciale n'est détecté selon les règles actuellement sélectionnées.",

  "helper.change.attribute.signals":
    "Des signes suggèrent que le processus a pu changer au cours de la période la plus récente du graphique ({chart}), notamment {signals}.",

  // Évaluation des changements : événements rares
  "helper.change.rare.stable":
    "Pour la période la plus récente du graphique ({chart}), aucun élément clair ne suggère un changement du processus. Aucun signal de cause spéciale n'est détecté selon les règles actuellement sélectionnées.",

  "helper.change.rare.signals":
    "Des signes suggèrent que le processus a pu changer au cours de la période la plus récente du graphique ({chart}), notamment {signals}.",

  "helper.change.investigate":
    "Ce signal invite à rechercher ce qui a changé dans le processus à cette période.",

  "helper.change.noChart":
    "Générez d'abord la carte actuellement sélectionnée pour que je puisse rechercher d'éventuels signes de changement.",

  // Règles applicables selon le type de carte
  "helper.rules.attribute":
    "Pour ce type de carte, SimpleSPC utilise les signaux de dépassement des limites de contrôle et la règle de séquence (décalage). Les règles de zone sont désactivées, car elles pourraient générer des alertes trompeuses sur les cartes aux attributs.",

  "helper.rules.rare":
    "Pour ce type de carte, les signaux de dépassement des limites de contrôle sont utilisés par défaut. Les règles de séquence et de tendance sont réservées aux options avancées, car les cartes d'événements rares peuvent naturellement présenter un comportement irrégulier.",

  "helper.rules.xmr":
    "Pour cette carte XmR, les signaux de dépassement des limites de contrôle et la règle de séquence sont utilisés par défaut. Les règles de tendance et les règles de zone sont proposées dans les options avancées, car l'ajout de règles supplémentaires peut augmenter le nombre de fausses alertes.",

  "helper.rules.xbars":
    "Pour cette carte X̄–S, les signaux de dépassement des limites de contrôle et la règle de séquence sont utilisés par défaut. Les règles de tendance et de zone sont disponibles dans les options avancées.",

  "helper.rules.run":
    "Pour cette carte de suivi (Run), SimpleSPC utilise les règles propres aux cartes de suivi, notamment la détection des décalages et, lorsqu'elle est activée, celle des tendances. Les règles de dépassement des limites de contrôle ne s'appliquent pas aux cartes de suivi.",

  "helper.rules.pareto":
    "Les règles de détection des signaux SPC ne s'appliquent pas aux diagrammes de Pareto. Un diagramme de Pareto classe les catégories selon leur fréquence plutôt que de représenter le comportement d'un processus au fil du temps.",

  "helper.rules.default":
    "SimpleSPC applique un ensemble prudent de règles adaptées à chaque type de carte. Toutes les règles ne sont pas disponibles pour tous les types de cartes.",

  //  MESSAGES DE REPLI DE L'ASSISTANT SPC

  "helper.response.empty":
    "Veuillez saisir une question sur le SPC ou sur votre graphique.",

  "helper.response.unsure":
    "Je ne suis pas certain de la réponse la plus adaptée. Voici quelques questions que vous pouvez essayer :",

  "helper.response.libraryMissing":
    "La bibliothèque de l'assistant SPC n'est pas disponible. Vérifiez que spc-helper-library.js est chargé avant spc.js.",

  // QUESTIONS PÉDAGOGIQUES SUR LES CARTES XmR

  "helper.q.xmrShortSignals":
    "Pourquoi les signaux XmR disparaissent-ils avec une courte période de référence ?",

  "helper.q.xmrMinimum":
    "Pourquoi une carte XmR nécessite-t-elle 12 points utilisables ?",

  "helper.q.xmrExclusionDisabled":
    "Pourquoi l'exclusion de ce point du calcul des limites est-elle désactivée ?",

  "helper.q.xmrNelsonOutside":
    "Pourquoi le filtrage de Nelson ne détecte-t-il pas les changements après la période de référence ?",

  // Aucune carte XmR disponible
  "helper.xmrEducation.noChart":
    "Générez une carte XmR pour obtenir des informations sur votre période actuelle. Voici l'explication générale.\n\n",

  // Période de référence courte et limites provisoires
  "helper.xmrEducation.short.intro":
    "SimpleSPC ne fournit volontairement pas d'interprétation formelle des signaux de cause spéciale lorsque les limites de contrôle XmR sont provisoires. Les limites peuvent rester affichées, mais un comportement inhabituel n'est pas présenté comme un signal SPC confirmé tant que la période utilisée pour estimer les limites ne contient pas suffisamment d'observations utilisables.",

  "helper.xmrEducation.short.provisional":
    "Votre période XmR la plus récente compte actuellement {count} points utilisables pour estimer les limites ; au moins {minimum} sont nécessaires pour l'interprétation. Cela explique pourquoi un signal peut disparaître lorsque vous raccourcissez la période de référence, même si les valeurs ultérieures semblent inhabituelles.",

  "helper.xmrEducation.short.provisionalUnknown":
    "La période XmR la plus récente ne satisfait pas encore au minimum de {minimum} points utilisables. Ses limites et son interprétation en matière de stabilité restent donc provisoires.",

  "helper.xmrEducation.short.assessable":
    "Votre période la plus récente compte {count} points utilisables et atteint le minimum de {minimum} points. Les règles SPC sélectionnées peuvent donc être interprétées. L'absence de signal ne signifie pas nécessairement que les performances sont satisfaisantes.",

  // Pourquoi douze points ?
  "helper.xmrEducation.minimum":
    "SimpleSPC exige, par prudence, au moins {minimum} points utilisables pour interpréter les limites XmR, les signaux de cause spéciale et la capabilité. Cette précaution évite de considérer comme suffisamment fiables des limites estimées à partir de très peu d'informations. Il s'agit d'une règle adoptée par cet outil, et non d'une garantie qu'un processus comportant 12 points est stable.\n\n" +
    "Cette distinction est importante : il ne suffit pas d'avoir {minimum} points quelque part sur le graphique. Ces points doivent contribuer à l'estimation des limites de la période concernée. Si la période de référence initiale ne contient que 10 points, ou si un changement de phase crée une nouvelle période courte, les limites peuvent rester provisoires. Les observations exclues ne sont pas automatiquement remplacées par des points ultérieurs.",

  "helper.xmrEducation.minimum.current":
    "\n\nPour votre période actuelle, {count} points utilisables contribuent à l'estimation des limites ; le minimum est de {minimum}.",

  // Exclusion manuelle indisponible
  "helper.xmrEducation.exclusion":
    "L'exclusion d'un point du calcul des limites de contrôle n'est utile que si ce point contribue effectivement aux limites estimées. Lorsque la période de référence initiale est fixée, un point ultérieur situé en dehors de cette période ne contribue pas au calcul. Son exclusion ne modifierait donc pas les limites. Une période créée par un changement de phase possède sa propre fenêtre d'estimation des limites.\n\n" +
    "SimpleSPC peut également désactiver une nouvelle exclusion si elle laisse trop peu d'observations incluses ou, dans le cas d'une carte XmR, aucune paire d'observations consécutives utilisables pour estimer la variation à partir des étendues mobiles.\n\n" +
    "Une observation exclue reste visible. Les étendues mobiles adjacentes sont omises de l'estimation de la variation. L'outil ne relie pas les observations de part et d'autre du point exclu pour créer une étendue mobile artificielle.\n\n" +
    "N'excluez une observation qu'après avoir identifié et étudié une cause spéciale connue. Si le processus semble avoir atteint durablement un nouveau niveau, examinez si un changement de phase justifié serait plus approprié. Le menu contextuel indique la raison précise lorsqu'une exclusion n'est pas disponible ; cette réponse de l'assistant n'identifie pas le point sur lequel vous avez cliqué.",

  // Filtrage de Nelson et détection des causes spéciales
  "helper.xmrEducation.nelson":
    "Le filtrage des étendues mobiles selon Nelson sert à estimer les limites de contrôle XmR. Il ne s'agit pas d'une règle destinée à signaler les changements sur l'ensemble du graphique.\n\n" +
    "SimpleSPC examine les étendues mobiles entre observations consécutives dans la période utilisée pour estimer les limites. Lorsque le filtrage est activé, les étendues mobiles strictement supérieures à 3,5 fois l'étendue mobile moyenne initiale sont omises de l'estimation de la variation. Les observations et les étendues mobiles d'origine restent visibles.\n\n" +
    "Un saut survenant après une période de référence fixée ne fait pas partie de l'estimation de la variation de cette période. Le filtrage de Nelson ne va donc ni le supprimer ni le signaler. Les observations ultérieures sont évaluées séparément selon les règles SPC sélectionnées, à condition que la période concernée possède suffisamment de données utilisables pour être interprétée. Un changement de phase crée une période distincte, avec sa propre estimation des limites.\n\n" +
    "Le filtrage de Nelson est également différent de l'exclusion manuelle d'un point : le filtrage omet certaines étendues mobiles de l'estimation de la variation, tandis qu'une exclusion manuelle retire du calcul une observation étudiée ainsi que ses étendues mobiles adjacentes.",

  "helper.xmrEducation.nelson.on":
    "\n\nSur votre carte XmR actuelle, le filtrage des étendues mobiles est activé.",

  "helper.xmrEducation.nelson.off":
    "\n\nSur votre carte XmR actuelle, le filtrage des étendues mobiles est désactivé.",

  // ASSISTANT DE CHOIX DU GRAPHIQUE

  "wizard.title": "Aidez-moi à choisir une carte",
  "wizard.subtitle": "Répondez à quelques questions pour obtenir une recommandation de carte et connaître les colonnes supplémentaires éventuellement nécessaires.",
  "wizard.close": "Fermer",
  "wizard.closeAria": "Fermer l'assistant de choix du graphique",

  "wizard.action.back": "Retour",
  "wizard.action.skip": "Passer",
  "wizard.action.manual": "Choisir manuellement",
  "wizard.action.use": "Utiliser cette carte",

  "wizard.start.question": "Que souhaitez-vous représenter ?",
  "wizard.start.option0": "Une mesure — par exemple un temps d'attente, un score, une température ou une durée de séjour",
  "wizard.start.option1": "Un nombre d'événements par période — par exemple des chutes par semaine ou des plaintes par mois",
  "wizard.start.option2": "Une proportion — par exemple 5 sur 100 ou un taux de conformité",
  "wizard.start.option3": "Des catégories — par exemple les causes de retard ou les types de plaintes",
  "wizard.start.option4": "Des événements rares — temps ou nombre d'opportunités entre les événements",
  "wizard.start.option5": "Je ne sais pas",

  "wizard.unsure.question": "Quelle description correspond le mieux à vos données ?",
  "wizard.unsure.hint": "Le type de carte dépend de ce que représente chaque valeur. Choisissez la description la plus proche de vos données. Si aucune ne convient, vous pouvez revenir à la liste et choisir une carte manuellement.",
  "wizard.unsure.option0": "Une mesure à chaque point temporel — par exemple un temps d'attente, un score ou une durée de séjour",
  "wizard.unsure.option1": "Un nombre d'événements par période — par exemple des chutes, des plaintes ou des incidents",
  "wizard.unsure.option2": "Une proportion ou un pourcentage — par exemple 5 sur 100 ou un taux de conformité",
  "wizard.unsure.option3": "Des catégories d'occurrences — par exemple les types de plaintes ou les motifs de retard",
  "wizard.unsure.option4": "Le temps ou le nombre d'opportunités entre des événements rares",

  "wizard.measurement.question": "Disposez-vous d'une seule mesure par point temporel ou de plusieurs mesures ?",
  "wizard.measurement.option0": "Une mesure à chaque point temporel",
  "wizard.measurement.option1": "Plusieurs mesures par point temporel — sous-groupes ou échantillons",
  "wizard.measurement.option2": "Je ne sais pas",

  "wizard.count.question": "Le volume d'activité ou le nombre d'opportunités varie-t-il ?",
  "wizard.count.hint": "Si le volume d'activité est globalement similaire à chaque période, considérez-le comme relativement constant. Si le volume varie fortement ou si vous disposez d'un dénominateur, par exemple des journées d'hospitalisation, des consultations ou des inspections, choisissez l'option indiquant une variation.<br><br>SimpleSPC commencera normalement par une carte U, puis vérifiera si une carte U′ de Laney correspond mieux à la variation observée.",
  "wizard.count.option0": "Non — volume globalement similaire à chaque période",
  "wizard.count.option1": "Oui — le volume varie ou je dispose d'une colonne de dénominateur",
  "wizard.count.option2": "Je ne sais pas",

  "wizard.proportion.question": "Disposez-vous des deux éléments nécessaires au calcul de la proportion ?",
  "wizard.proportion.hint": "Par exemple, si 8 patients sur 100 satisfont au critère, le numérateur est 8 et le dénominateur est 100.<br><br>SimpleSPC commencera normalement par une carte P, puis vérifiera si une carte P′ de Laney correspond mieux à la variation observée.",
  "wizard.proportion.option0": "Oui — j'ai des colonnes pour le numérateur et le dénominateur",
  "wizard.proportion.option1": "Non — je dispose uniquement du pourcentage ou de la proportion",
  "wizard.proportion.option2": "Je ne sais pas",

  "wizard.pareto.question": "Comment vos données de catégories sont-elles enregistrées ?",
  "wizard.pareto.hint": "Un diagramme de Pareto compare la fréquence des différentes catégories. Vous pouvez utiliser des enregistrements individuels ou des données contenant déjà les effectifs de chaque catégorie.",
  "wizard.pareto.option0": "Une ligne par occurrence — par exemple une ligne pour chaque plainte",
  "wizard.pareto.option1": "Effectifs déjà regroupés — par exemple Chutes : 20, Retards : 15",
  "wizard.pareto.option2": "Je ne sais pas",

  "wizard.rare.question": "Quelle description correspond le mieux à vos données ?",
  "wizard.rare.hint": "Choisissez cette option lorsque l'événement est peu fréquent et que vous examinez l'intervalle entre les événements.",
  "wizard.rare.option0": "Temps entre les événements — par exemple le nombre de jours entre des incidents",
  "wizard.rare.option1": "Opportunités entre les événements — par exemple le nombre d'interventions entre des incidents",
  "wizard.rare.option2": "Je ne sais pas",

  "wizard.result.heading": "Carte recommandée :",
  "wizard.result.hint": "Vous pouvez toujours choisir manuellement un autre type de carte.",

  "wizard.chart.xmr": "Carte XmR",
  "wizard.chart.xbars": "Carte X̄–S",
  "wizard.chart.c": "Carte C",
  "wizard.chart.p": "Carte P",
  "wizard.chart.u": "Carte U",
  "wizard.chart.t": "Carte T",
  "wizard.chart.g": "Carte G",
  "wizard.chart.pareto": "Diagramme de Pareto",

  "wizard.reason.measurementSingle": "Vous disposez d'une mesure par point temporel.",
  "wizard.reason.measurementSubgroups": "Vous disposez de plusieurs mesures par point temporel, regroupées en sous-groupes.",
  "wizard.reason.countConstant": "Vous disposez de nombres d'événements avec un volume d'activité ou un nombre d'opportunités relativement constant.",
  "wizard.reason.countVaries": "Vous disposez de nombres d'événements dont le volume d'activité ou le nombre d'opportunités varie. Une carte U constitue donc un bon point de départ. Après la génération, SimpleSPC vérifiera si les taux varient davantage ou moins que ne le prévoit une carte U standard. Si nécessaire, il proposera une carte U′ de Laney et en expliquera la raison.",
  "wizard.reason.proportionYes": "Vous disposez d'un numérateur et d'un dénominateur pour chaque point temporel. Une carte P constitue donc un bon point de départ. Après la génération, SimpleSPC vérifiera si les données varient davantage ou moins que ne le prévoit une carte P standard. Si nécessaire, il proposera une carte P′ de Laney et en expliquera la raison.",
  "wizard.reason.proportionNo": "Sans numérateur et dénominateur, une carte P ou P′ ne peut pas calculer correctement la variation attendue. Vous pouvez plutôt représenter le pourcentage ou la proportion comme une mesure unique au fil du temps, à l'aide d'une carte XmR.",
  "wizard.reason.paretoCounts": "Vous souhaitez comparer des catégories et déterminer celles qui regroupent le plus d'occurrences. Un diagramme de Pareto les classe par fréquence et présente leur pourcentage cumulé. Vos données contiennent déjà les effectifs : sélectionnez la colonne des catégories et la colonne numérique des fréquences. Le diagramme aide à définir les priorités d'investigation, mais ne permet pas d'identifier les causes profondes ni d'évaluer la stabilité statistique du processus.",
  "wizard.reason.paretoRows": "Vous souhaitez comparer des catégories et déterminer celles qui regroupent le plus d'occurrences. Un diagramme de Pareto les classe par fréquence et présente leur pourcentage cumulé. Si chaque ligne correspond à une occurrence, sélectionnez la colonne des catégories et l'option permettant de compter chaque ligne comme une occurrence. Le diagramme aide à définir les priorités d'investigation, mais ne permet pas d'identifier les causes profondes ni d'évaluer la stabilité statistique du processus.",
  "wizard.reason.rareTime": "Vous disposez du temps entre des événements rares.",
  "wizard.reason.rareOpportunities": "Vous disposez du nombre d'opportunités entre des événements rares.",
  "wizard.reason.unsure": "En cas d'incertitude, une carte XmR constitue un point de départ approprié pour une mesure unique suivie au fil du temps.",

  // LOT 36A-2 — FENÊTRE DE PARAMÉTRAGE DES CARTES

  "setup.title": "Paramètres du graphique",
  "setup.close": "Fermer",
  "setup.closeAria": "Fermer les paramètres du graphique",
  "setup.dontShow": "Ne plus afficher automatiquement cette fenêtre",

  // Sous-titres selon le type de carte
  "setup.subtitle.run": "Carte de suivi (Run) : évolution des valeurs au fil du temps.",
  "setup.subtitle.xmr": "Carte XmR : mesures individuelles suivies au fil du temps.",
  "setup.subtitle.c": "Carte C : nombre d'événements par période.",
  "setup.subtitle.p": "Carte P : proportion par rapport à un total.",
  "setup.subtitle.pprime": "Carte P′ de Laney : carte de proportions avec limites ajustées.",
  "setup.subtitle.u": "Carte U : taux par opportunité.",
  "setup.subtitle.uprime": "Carte U′ de Laney : carte de taux avec limites ajustées.",
  "setup.subtitle.xbars": "Carte X̄–S : mesures regroupées en sous-groupes.",
  "setup.subtitle.t": "Carte T : temps entre des événements rares.",
  "setup.subtitle.g": "Carte G : opportunités entre des événements rares.",
  "setup.subtitle.pareto": "Diagramme de Pareto : comparer les catégories et définir des priorités.",

  // Titres des rubriques
  "setup.heading.when": "À utiliser lorsque…",
  "setup.heading.columns": "Colonnes nécessaires",
  "setup.heading.example": "Exemple de structure des données",
  "setup.heading.mistake": "Erreur fréquente",
  "setup.heading.mistakes": "Erreurs fréquentes",
  "setup.heading.pprime": "Choisir entre P et P′",
  "setup.heading.uprime": "Choisir entre U et U′",
  "setup.heading.rates": "Affichage des taux",
  "setup.heading.paretoRows": "Exemple 1 : une ligne par occurrence",
  "setup.heading.paretoCounts": "Exemple 2 : effectifs déjà regroupés",
  "setup.heading.paretoShows": "Ce que montre le diagramme",
  "setup.heading.tMode": "Choisissez votre mode de saisie",
  "setup.heading.fallback": "Conseils de paramétrage",

  // En-têtes et catégories des exemples
  "setup.example.week": "Semaine",
  "setup.example.waitingTime": "Temps d'attente",
  "setup.example.date": "Date",
  "setup.example.lengthStay": "Durée de séjour",
  "setup.example.falls": "Chutes",
  "setup.example.patientsHarm": "Patients ayant subi un préjudice",
  "setup.example.patientsReviewed": "Patients évalués",
  "setup.example.month": "Mois",
  "setup.example.infections": "Infections",
  "setup.example.bedDays": "Journées d'hospitalisation",
  "setup.example.jan": "Janv.",
  "setup.example.feb": "Févr.",
  "setup.example.mar": "Mars",
  "setup.example.day": "Jour",
  "setup.example.reading": "Mesure",
  "setup.example.sampleId": "ID_sous-groupe",
  "setup.example.mon": "Lun.",
  "setup.example.tue": "Mar.",
  "setup.example.eventDate": "Date de l'événement",
  "setup.example.procedures": "Interventions entre événements indésirables",
  "setup.example.category": "Catégorie",
  "setup.example.count": "Effectif",
  "setup.example.medication": "Médicaments",
  "setup.example.pressureDamage": "Lésions de pression",

  //  CONTENU DES CARTES STANDARD

  // RUN
  "setup.run.card1": `
    <p style="margin:0;">Vous souhaitez suivre une mesure au fil du temps à l'aide d'une ligne centrale correspondant à la <strong>médiane</strong>.</p>
  `,
  "setup.run.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → temps ou ordre des observations</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → mesure à suivre</li>
    </ul>
  `,
  "setup.run.card4": `
    <p style="margin:0;">Ne sélectionnez pas un numéro de ligne ou une colonne d'identifiants comme mesure. Le graphique pourrait sembler correct sans avoir de signification statistique.</p>
  `,

  // XmR
  "setup.xmr.card1": `
    <p style="margin:0;">Vous disposez d'<strong>une mesure par point temporel</strong> et souhaitez obtenir une moyenne et des limites de contrôle.</p>
  `,
  "setup.xmr.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → date, semaine, mois ou séquence</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → mesure numérique</li>
    </ul>
  `,
  "setup.xmr.card4": `
    <p style="margin:0;">Utilisez XmR lorsque vous disposez d'une seule mesure par point temporel. Si chaque point temporel comporte plusieurs mesures, utilisez plutôt une carte <strong>X̄–S</strong>.</p>
  `,

  // C
  "setup.c.card1": `
    <p style="margin:0;">Vous représentez un <strong>nombre d'événements</strong> par période, avec un volume d'opportunités relativement constant.</p>
  `,
  "setup.c.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → date, semaine, mois ou séquence</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → nombre d'événements</li>
    </ul>
  `,
  "setup.c.card4": `
    <p style="margin:0;">Si le dénominateur varie fortement d'une période à l'autre, une <strong>carte U</strong> peut être plus adaptée qu'une carte C.</p>
  `,

  // P
  "setup.p.card1": `
    <p style="margin:0;">Vous souhaitez suivre une <strong>proportion</strong>, par exemple 5 sur 100 ou le pourcentage de cas respectant une norme.</p>
  `,
  "setup.p.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → temps ou ordre des observations</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → numérateur</li>
      <li><strong>Troisième colonne</strong> → dénominateur</li>
    </ul>
  `,
  "setup.p.card4": `
    <p style="margin:0;">Le numérateur ne doit pas être supérieur au dénominateur.</p>
  `,

  // U
  "setup.u.card1": `
    <p style="margin:0;">Vous souhaitez suivre un <strong>taux</strong> lorsque le dénominateur varie d'un point à l'autre.</p>
  `,
  "setup.u.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → temps ou ordre des observations</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → nombre d'événements</li>
      <li><strong>Troisième colonne</strong> → opportunités ou exposition</li>
    </ul>
  `,
  "setup.u.card4": `
    <p style="margin:0;">Utilisez une carte U lorsque le dénominateur varie. S'il reste relativement constant, une <strong>carte C</strong> peut être plus adaptée.</p>
  `,

  // X̄–S
  "setup.xbars.card1": `
    <p style="margin:0;">Vous disposez de <strong>plusieurs mesures dans chaque sous-groupe</strong> et souhaitez étudier à la fois les moyennes des sous-groupes et la variation à l'intérieur de ceux-ci.</p>
  `,
  "setup.xbars.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → libellé de sous-groupe ou repère temporel</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → valeur mesurée</li>
      <li><strong>Troisième colonne</strong> → identifiant du sous-groupe</li>
    </ul>
  `,
  "setup.xbars.card4": `
    <p style="margin:0;">N'utilisez pas X̄–S si chaque sous-groupe ne contient qu'une seule mesure. Utilisez plutôt une carte <strong>XmR</strong>.</p>
  `,

  // LANEY P′
  "setup.pprime.card1": `
    <p style="margin:0;">Vous suivez une <strong>proportion</strong> à partir d'un numérateur et d'un dénominateur, mais la variation observée est plus forte ou plus faible que celle attendue avec une carte P standard. La carte P′ ajuste la largeur des limites de contrôle pour tenir compte de cette variation supplémentaire ou réduite.</p>
  `,
  "setup.pprime.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → temps ou ordre des observations</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → numérateur (nombre de cas répondant à la définition)</li>
      <li><strong>Troisième colonne</strong> → dénominateur (nombre total de cas évalués)</li>
    </ul>
  `,
  "setup.pprime.card4": `
    <p style="margin:0;">
      Commencez normalement par une <strong>carte P</strong>. SimpleSPC vérifie si la variation observée diffère de celle attendue avec une carte P standard et peut recommander de passer à une carte P′.
      <br><br>
      Le facteur d'ajustement est appelé <strong>Sigma Z</strong>. Il modifie les limites de contrôle, mais pas les proportions représentées. Le choix de changer de carte vous appartient.
    </p>
  `,
  "setup.pprime.card5": `
    <p style="margin:0;">Saisissez les effectifs réels du numérateur et du dénominateur, et non un pourcentage déjà calculé. Le numérateur ne peut pas dépasser le dénominateur.</p>
  `,

  // LANEY U′
  "setup.uprime.card1": `
    <p style="margin:0;">Vous suivez le <strong>nombre d'événements par opportunité</strong>, mais la variation observée est plus forte ou plus faible que celle attendue avec une carte U standard. La carte U′ ajuste les limites de contrôle pour refléter cette variation.</p>
  `,
  "setup.uprime.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → temps ou ordre des observations</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → nombre d'événements</li>
      <li><strong>Troisième colonne</strong> → nombre d'opportunités ou niveau d'exposition</li>
    </ul>
  `,
  "setup.uprime.card4": `
    <p style="margin:0;">
      Commencez normalement par une <strong>carte U</strong>. SimpleSPC vérifie si la variation observée diffère de celle attendue avec une carte U standard et peut recommander de passer à une carte U′.
      <br><br>
      Le facteur d'ajustement est appelé <strong>Sigma Z</strong>. Il modifie les limites de contrôle, mais pas les taux d'événements sous-jacents. Le choix de changer de carte vous appartient.
    </p>
  `,
  "setup.uprime.card5": `
    <p style="margin:0;">Dans l'onglet Graphique, utilisez l'option <strong>Afficher le taux pour</strong> afin de présenter les taux pour 1, 100, 1 000 ou une autre unité disponible. Cette option modifie l'échelle d'affichage, mais pas le calcul statistique.</p>
  `,
  "setup.uprime.card6": `
    <p style="margin:0;">Saisissez les nombres réels d'événements et d'opportunités, et non un taux déjà calculé. Contrairement à une proportion, un taux peut dépasser un événement par opportunité.</p>
  `,

  // T — ÉVÉNEMENTS RARES
  "setup.t.card1": `
    <p style="margin:0;">Vous souhaitez suivre le <strong>temps entre des événements rares</strong>.</p>
  `,
  "setup.t.card2": `
    <label style="display:block; margin:0.25rem 0;">
      <input type="radio" name="tChartInputMode" value="eventDates" {checkedDates}>
      <strong>Je dispose des dates des événements</strong> (une ligne par événement)
    </label>
    <div class="hint small-hint" style="margin-top:0.15rem; margin-bottom:0.5rem;">
      Indiquez la date ou l'heure de l'événement dans la <em>colonne des dates / axe X</em>. La <em>colonne des valeurs / axe Y</em> n'est pas utilisée.
    </div>
    <label style="display:block; margin:0.25rem 0;">
      <input type="radio" name="tChartInputMode" value="gaps" {checkedGaps}>
      <strong>Je dispose déjà des intervalles</strong> (temps numérique entre événements)
    </label>
    <div class="hint small-hint" style="margin-top:0.15rem;">
      Saisissez les valeurs des intervalles dans la <em>colonne des valeurs / axe Y</em>.
    </div>
  `,
  "setup.t.card4": `
    <p style="margin:0;">N'utilisez pas une carte T pour représenter un nombre d'événements par mois. Utilisez plutôt une carte <strong>C</strong>, <strong>P</strong> ou <strong>U</strong>.</p>
  `,

  // G — ÉVÉNEMENTS RARES
  "setup.g.card1": `
    <p style="margin:0;">Vous souhaitez suivre le <strong>nombre d'opportunités entre des événements rares</strong>.</p>
  `,
  "setup.g.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des dates / axe X</strong> → colonne facultative de libellé ou d'ordre</li>
      <li><strong>Colonne des valeurs / axe Y</strong> → nombre d'opportunités entre les événements</li>
    </ul>
  `,
  "setup.g.card4": `
    <p style="margin:0;">Utilisez une carte G si l'intervalle entre les événements correspond à un <strong>nombre d'opportunités</strong>. S'il s'agit d'une durée écoulée, utilisez une <strong>carte T</strong>.</p>
  `,

  // PARETO
  "setup.pareto.card1": `
    <p style="margin:0;">Vous souhaitez déterminer les catégories qui regroupent le plus grand nombre d'occurrences enregistrées. Un diagramme de Pareto classe les catégories de la plus fréquente à la moins fréquente et présente une courbe de pourcentage cumulé.</p>
  `,
  "setup.pareto.card2": `
    <ul style="margin:0;">
      <li><strong>Colonne des catégories</strong> → catégorie à comparer</li>
      <li><strong>Colonne des fréquences / effectifs</strong> → méthode de dénombrement des occurrences</li>
    </ul>
    <p style="margin-bottom:0;">Si chaque ligne correspond à une occurrence, choisissez <strong>(compter chaque ligne comme 1)</strong>. Si les données sont déjà regroupées, sélectionnez la colonne numérique contenant les effectifs.</p>
  `,
  "setup.pareto.card5": `
    <p style="margin:0;">
      Les barres bleues indiquent le nombre d'occurrences dans chaque catégorie. La courbe rouge représente le pourcentage cumulé.
      <br><br>
      La ligne pointillée des 80 % sert à définir des priorités : ce n'est <strong>pas une limite de contrôle statistique</strong>.
    </p>
  `,
  "setup.pareto.card6": `
    <p style="margin:0;">
      Ne supposez pas que les catégories les plus fréquentes correspondent nécessairement aux causes sous-jacentes du problème. Elles constituent des points de départ utiles pour les investigations.
      <br><br>
      Lorsque vous utilisez une colonne d'effectifs numériques, les valeurs doivent être des nombres entiers non négatifs. Il faut au moins deux catégories présentant des effectifs positifs.
    </p>
  `,

//ENREGISTREMENT DU GRAPHIQUE PORTABLE

"portable.saveConfirm":
  "Enregistrer une copie hors ligne modifiable de cette carte SPC ?\n\n" +
  "Le fichier HTML contiendra les données du graphique ainsi que " +
  "l'application FreeSPC. Conservez et partagez ce fichier conformément " +
  "aux exigences de gouvernance de l'information de votre organisation.",

"portable.creating":
  "Création du fichier portable…",

//  XmR detailed-summary labels.
  "summary.xmr.detail.above": "Au-dessus",
  "summary.xmr.detail.below": "Au-dessous",
  "summary.xmr.detail.increasing": "Croissante",
  "summary.xmr.detail.decreasing": "Décroissante",
  "summary.xmr.detail.disabled": "Non évalué (règle désactivée)",
  "summary.xmr.detail.distinct": "Observations distinctes signalées",
  "summary.xmr.detail.overlap": "Une observation peut satisfaire plusieurs règles.",
  "summary.xmr.detail.beyond": "Hors des limites de contrôle",
  "summary.xmr.detail.zone23": "2 points sur 3 au-delà de 2σ",
  "summary.xmr.detail.zone45": "4 points sur 5 au-delà de 1σ",
  "summary.xmr.detail.windows": "fenêtres répondant à la règle",
  "summary.xmr.detail.stretches": "séquences regroupées",
  "summary.xmr.detail.outliers": "Autres observations inhabituelles",

  // Novembre 2026 — Nouveautés
  "whatsNew.button": "Nouveautés",
  "whatsNew.bannerTitle": "SimpleSPC a été mis à jour !",
  "whatsNew.bannerText": "De nouvelles langues et des explications plus claires sont disponibles. Les fonctionnalités habituelles sont toujours là.",
  "whatsNew.view": "Voir les nouveautés",
  "whatsNew.dismiss": "Masquer",
  "whatsNew.title": "Quoi de neuf dans SimpleSPC ?",
  "whatsNew.release": "Mise à jour de novembre 2026 · v2026.11",
  "whatsNew.intro": "Nous avons apporté quelques améliorations pour faciliter l’utilisation et la compréhension des cartes SPC.",
  "whatsNew.languagesTitle": "Plus de langues",
  "whatsNew.languagesText": "Choisissez parmi six langues, avec une aide et des explications de graphiques traduites.",
  "whatsNew.signalsTitle": "Explications plus claires",
  "whatsNew.signalsText": "Découvrez les observations et les règles à l’origine des signaux de cause spéciale.",
  "whatsNew.refinementsTitle": "Améliorations et corrections",
  "whatsNew.refinementsText": "Améliorations des infobulles, des cartes X̄–S, des formats numériques et des exportations.",
  "whatsNew.reassure": "Vous pouvez continuer à utiliser SimpleSPC comme avant. Vous n’avez rien à faire différemment.",
  "whatsNew.history": "Versions précédentes sur GitHub",
  "whatsNew.closeAria": "Fermer les nouveautés",

//Button tooltips
  "buttons.helpTitle": "Ouvrir les instructions et les conseils pratiques pour utiliser SimpleSPC.",
  "buttons.spcHelperTitle": "Poser des questions sur le choix, la création et l’interprétation des cartes SPC.",
  "whatsNew.buttonTitle": "Découvrir les changements apportés lors de la dernière mise à jour de SimpleSPC."




});
