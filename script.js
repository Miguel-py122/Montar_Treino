const STORAGE_KEY = "planner-treinos:v3";
const LEGACY_STORAGE_KEYS = [STORAGE_KEY, "planner-treinos:v2", "planner-treinos:v1"];
const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const EXERCISE_LIBRARY = Object.freeze({
  Peito: [
    "Supino reto",
    "Supino inclinado",
    "Supino declinado",
    "Crucifixo reto",
    "Crucifixo inclinado",
    "Cross over",
    "Peck deck",
    "Flexao de bracos",
  ],
  Costas: [
    "Puxada frontal",
    "Puxada supinada",
    "Remada curvada",
    "Remada baixa",
    "Remada serrote",
    "Remada cavalinho",
    "Pulldown",
    "Barra fixa",
    "Levantamento terra",
    "Pull over",
  ],
  Pernas: [
    "Agachamento livre",
    "Agachamento hack",
    "Leg press",
    "Cadeira extensora",
    "Mesa flexora",
    "Stiff",
    "Afundo",
    "Passada",
    "Agachamento bulgaro",
    "Terra romeno",
  ],
  Gluteos: [
    "Elevacao pelvica",
    "Ponte glutea",
    "Coice no cabo",
    "Abducao de quadril",
    "Agachamento sumo",
    "Step-up",
  ],
  Ombros: [
    "Desenvolvimento com halteres",
    "Desenvolvimento militar",
    "Arnold press",
    "Elevacao lateral",
    "Elevacao frontal",
    "Crucifixo inverso",
    "Face pull",
    "Remada alta",
  ],
  Biceps: [
    "Rosca direta",
    "Rosca alternada",
    "Rosca martelo",
    "Rosca Scott",
    "Rosca concentrada",
    "Rosca inversa",
  ],
  Triceps: [
    "Triceps testa",
    "Triceps corda",
    "Triceps frances",
    "Supino fechado",
    "Mergulho",
    "Triceps banco",
  ],
  Abdomen: [
    "Prancha",
    "Crunch",
    "Abdominal infra",
    "Abdominal na polia",
    "Elevacao de pernas",
    "Ab wheel",
  ],
  Panturrilhas: [
    "Panturrilha em pe",
    "Panturrilha sentado",
    "Panturrilha no leg press",
    "Panturrilha unilateral",
    "Donkey calf raise",
  ],
});

const SPLIT_PRESET_LIBRARY = Object.freeze({
  Push: {
    groups: ["Peito", "Ombros", "Triceps"],
    exercises: [
      { category: "Peito", name: "Supino reto", sets: "4", reps: "8", load: "" },
      { category: "Peito", name: "Cross over", sets: "3", reps: "12", load: "" },
      { category: "Ombros", name: "Desenvolvimento com halteres", sets: "4", reps: "10", load: "" },
      { category: "Ombros", name: "Elevacao lateral", sets: "3", reps: "12", load: "" },
      { category: "Triceps", name: "Triceps corda", sets: "3", reps: "12", load: "" },
    ],
  },
  Pull: {
    groups: ["Costas", "Biceps"],
    exercises: [
      { category: "Costas", name: "Puxada frontal", sets: "4", reps: "10", load: "" },
      { category: "Costas", name: "Remada baixa", sets: "4", reps: "10", load: "" },
      { category: "Costas", name: "Barra fixa", sets: "3", reps: "8", load: "" },
      { category: "Biceps", name: "Rosca direta", sets: "3", reps: "10", load: "" },
      { category: "Biceps", name: "Rosca martelo", sets: "3", reps: "12", load: "" },
    ],
  },
  Legs: {
    groups: ["Pernas", "Gluteos", "Panturrilhas"],
    exercises: [
      { category: "Pernas", name: "Agachamento livre", sets: "4", reps: "8", load: "" },
      { category: "Pernas", name: "Leg press", sets: "4", reps: "12", load: "" },
      { category: "Pernas", name: "Mesa flexora", sets: "3", reps: "12", load: "" },
      { category: "Gluteos", name: "Elevacao pelvica", sets: "4", reps: "10", load: "" },
      { category: "Panturrilhas", name: "Panturrilha em pe", sets: "4", reps: "15", load: "" },
    ],
  },
  Upper: {
    groups: ["Peito", "Costas", "Ombros", "Biceps", "Triceps"],
    exercises: [
      { category: "Peito", name: "Supino inclinado", sets: "4", reps: "8", load: "" },
      { category: "Costas", name: "Remada curvada", sets: "4", reps: "8", load: "" },
      { category: "Ombros", name: "Arnold press", sets: "3", reps: "10", load: "" },
      { category: "Biceps", name: "Rosca alternada", sets: "3", reps: "10", load: "" },
      { category: "Triceps", name: "Triceps testa", sets: "3", reps: "10", load: "" },
    ],
  },
  Lower: {
    groups: ["Pernas", "Gluteos", "Panturrilhas", "Abdomen"],
    exercises: [
      { category: "Pernas", name: "Agachamento hack", sets: "4", reps: "10", load: "" },
      { category: "Pernas", name: "Stiff", sets: "4", reps: "10", load: "" },
      { category: "Gluteos", name: "Coice no cabo", sets: "3", reps: "12", load: "" },
      { category: "Panturrilhas", name: "Panturrilha sentado", sets: "4", reps: "15", load: "" },
      { category: "Abdomen", name: "Prancha", sets: "3", reps: "30", load: "" },
    ],
  },
  Core: {
    groups: ["Abdomen", "Ombros"],
    exercises: [
      { category: "Abdomen", name: "Crunch", sets: "3", reps: "15", load: "" },
      { category: "Abdomen", name: "Elevacao de pernas", sets: "3", reps: "12", load: "" },
      { category: "Abdomen", name: "Ab wheel", sets: "3", reps: "10", load: "" },
      { category: "Ombros", name: "Face pull", sets: "3", reps: "15", load: "" },
      { category: "Ombros", name: "Crucifixo inverso", sets: "3", reps: "12", load: "" },
    ],
  },
});

const ALL_EXERCISES = Object.values(EXERCISE_LIBRARY).flat();
const EXERCISE_CATEGORY_BY_NAME = new Map(
  Object.entries(EXERCISE_LIBRARY).flatMap(([category, exercises]) =>
    exercises.map((exercise) => [normalizeLookupValue(exercise), category]),
  ),
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
  bodyMapSelectedLabel: document.querySelector("#body-map-selected-label"),
  bodyMapSelectedCopy: document.querySelector("#body-map-selected-copy"),
  exerciseDrawer: document.querySelector("#exercise-drawer"),
  exerciseDrawerTraining: document.querySelector("#exercise-drawer-training"),
  exerciseDrawerCategory: document.querySelector("#exercise-drawer-category"),
  exerciseDrawerDescription: document.querySelector("#exercise-drawer-description"),
  exerciseDrawerSearch: document.querySelector("#exercise-drawer-search"),
  exerciseDrawerMeta: document.querySelector("#exercise-drawer-meta"),
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
    { id: "training-a", name: "Treino A", splitLabel: "", focusGroups: [], bodyMapCategory: "", bodyMapView: "front", exercises: [defaultExercise()] },
    { id: "training-b", name: "Treino B", splitLabel: "", focusGroups: [], bodyMapCategory: "", bodyMapView: "front", exercises: [] },
    { id: "training-c", name: "Treino C", splitLabel: "", focusGroups: [], bodyMapCategory: "", bodyMapView: "front", exercises: [] },
    { id: "training-d", name: "Treino D", splitLabel: "", focusGroups: [], bodyMapCategory: "", bodyMapView: "front", exercises: [] },
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
        bodyMapCategory: training.bodyMapCategory || "",
        bodyMapView: training.bodyMapView === "back" ? "back" : "front",
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
  const activeGroups = new Set(activeTraining.focusGroups || []);
  const selectedCategory = activeTraining.bodyMapCategory || activeTraining.focusGroups[0] || "";

  elements.activeSplitLabel.textContent = activeTraining.splitLabel
    ? `${activeTraining.splitLabel} • ${activeTraining.focusGroups.join(", ")}`
    : "Sem preset aplicado.";

  elements.splitPresets.querySelectorAll("[data-action='apply-split']").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.split === activeTraining.splitLabel);
  });

  document.querySelectorAll(".body-hotspot").forEach((button) => {
    const category = button.dataset.category || "";
    button.classList.toggle("is-active", category === selectedCategory || activeGroups.has(category));
  });

  if (elements.bodyMapSelectedLabel) {
    elements.bodyMapSelectedLabel.textContent = selectedCategory || "Nenhum grupo selecionado";
  }

  if (elements.bodyMapSelectedCopy) {
    elements.bodyMapSelectedCopy.textContent = selectedCategory
      ? `${selectedCategory} esta destacado no mapa. Clique novamente em outra regiao para trocar rapidamente o foco do treino.`
      : "Escolha uma regiao do corpo para o sistema aplicar esse agrupamento na proxima linha livre.";
  }

  setBodyMapView(activeTraining.bodyMapView || "front", false);
}

function setBodyMapView(view, shouldPersist = true) {
  const activeTraining = getActiveTraining();
  activeTraining.bodyMapView = view === "back" ? "back" : "front";

  document.querySelectorAll('[data-action="toggle-map-view"]').forEach((button) => {
    const isActive = button.dataset.view === activeTraining.bodyMapView;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });

  document.querySelectorAll('[data-view-panel]').forEach((panel) => {
    const isActive = panel.dataset.viewPanel === activeTraining.bodyMapView;
    panel.hidden = !isActive;
    panel.classList.toggle("is-active", isActive);
  });

  if (shouldPersist) {
    persistAndRender(false);
  }
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
    bodyMapCategory: "",
    bodyMapView: "front",
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

  if (target instanceof HTMLElement) {
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

    const splitButton = target.closest('[data-action="apply-split"]');
    if (splitButton) {
      applySplitPreset(splitButton.dataset.split || "", (splitButton.dataset.groups || "").split(",").filter(Boolean));
      return;
    }

    const viewButton = target.closest('[data-action="toggle-map-view"]');
    if (viewButton) {
      setBodyMapView(viewButton.dataset.view || "front");
      return;
    }

    const bodyZoneButton = target.closest('[data-action="quick-category"]');
    if (bodyZoneButton) {
      assignCategoryFromBodyMap(bodyZoneButton.dataset.category || "");
      return;
    }
  }

  if (target instanceof HTMLElement && (target.closest(".exercise-picker") || target.closest(".category-picker") || target.closest(".exercise-drawer__panel") || target.closest(".category-drawer__panel"))) {
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
  const source = category && EXERCISE_LIBRARY[category] ? EXERCISE_LIBRARY[category] : ALL_EXERCISES;
  const normalizedQuery = normalizeLookupValue(query);

  if (!normalizedQuery) {
    return source;
  }

  const startsWithMatches = source.filter((exercise) => normalizeLookupValue(exercise).startsWith(normalizedQuery));
  const containsMatches = source.filter(
    (exercise) =>
      !normalizeLookupValue(exercise).startsWith(normalizedQuery)
      && normalizeLookupValue(exercise).includes(normalizedQuery),
  );

  return [...startsWithMatches, ...containsMatches];
}

function openExerciseDrawer(row, exercise) {
  if (!(row instanceof HTMLTableRowElement) || !exercise || !elements.exerciseDrawer) {
    return;
  }

  exerciseDrawerState.exerciseId = exercise.id;
  exerciseDrawerState.query = exercise.name || "";
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

  if (elements.exerciseDrawerTraining) {
    elements.exerciseDrawerTraining.textContent = getDisplayTrainingName(activeTraining.name);
  }

  if (elements.exerciseDrawerCategory) {
    elements.exerciseDrawerCategory.textContent = categoryLabel;
  }

  if (elements.exerciseDrawerDescription) {
    elements.exerciseDrawerDescription.textContent = exercise.category
      ? `Exibindo exercicios de ${exercise.category}. Ajuste a categoria da linha para trocar o agrupamento.`
      : "Sem categoria definida. Escolha qualquer exercicio e a categoria sera ajustada automaticamente.";
  }

  if (elements.exerciseDrawerSearch && elements.exerciseDrawerSearch.value !== exerciseDrawerState.query) {
    elements.exerciseDrawerSearch.value = exerciseDrawerState.query;
  }

  if (elements.exerciseDrawerMeta) {
    elements.exerciseDrawerMeta.textContent = `${categoryLabel} • ${matches.length} opcao(oes)`;
  }

  if (!elements.exerciseDrawerList) {
    return;
  }

  if (!visibleMatches.length) {
    elements.exerciseDrawerList.innerHTML = '<div class="exercise-option exercise-option--empty">Nenhum exercicio encontrado.</div>';
    return;
  }

  elements.exerciseDrawerList.innerHTML = visibleMatches
    .map((exerciseName) => {
      const inferredCategory = inferCategoryFromExerciseName(exerciseName);
      const isSelected = normalizeLookupValue(exerciseName) === normalizedSelectedName;
      return `<button class="exercise-option${isSelected ? " is-selected" : ""}" type="button" data-action="choose-exercise-drawer" data-exercise="${escapeAttribute(exerciseName)}" data-category="${escapeAttribute(inferredCategory)}"><span>${exerciseName}</span><small>${inferredCategory || categoryLabel}</small></button>`;
    })
    .join("");
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
      return `<button class="category-option category-option--drawer${isActive ? " is-active" : ""}" type="button" data-action="choose-category-drawer" data-category="${category}"><span>${category}</span><small>${exerciseCount} exercicio${exerciseCount === 1 ? "" : "s"}</small></button>`;
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
  const insertedCount = preset ? seedExercisesFromPreset(activeTraining, preset.exercises) : 0;

  activeTraining.splitLabel = splitLabel;
  activeTraining.focusGroups = presetGroups;
  activeTraining.bodyMapCategory = presetGroups[0] || "";
  persistAndRender();
  setStatusMessage(`Preset ${splitLabel} aplicado com ${insertedCount} exercicio(s) base.`, "success");
}

function assignCategoryFromBodyMap(category) {
  const activeTraining = getActiveTraining();
  let exercise = activeTraining.exercises.find((item) => !isExerciseMeaningful(item));

  if (!exercise) {
    exercise = defaultExercise();
    activeTraining.exercises.push(exercise);
  }

  exercise.category = category;
  activeTraining.bodyMapCategory = category;

  if (!activeTraining.focusGroups.includes(category)) {
    activeTraining.focusGroups = [...activeTraining.focusGroups, category];
  }

  persistAndRender();
  setStatusMessage(`${category} aplicado na proxima linha disponivel e destacado no mapa corporal.`, "success");
}

function seedExercisesFromPreset(training, presetExercises) {
  const meaningfulExercises = training.exercises.filter(isExerciseMeaningful);
  const existingNames = new Set(meaningfulExercises.map((exercise) => normalizeLookupValue(exercise.name)));
  const exercisesToInsert = presetExercises.filter(
    (exercise) => !existingNames.has(normalizeLookupValue(exercise.name)),
  );

  training.exercises = meaningfulExercises;
  exercisesToInsert.forEach((exercise) => {
    training.exercises.push({
      id: crypto.randomUUID(),
      category: exercise.category,
      name: exercise.name,
      sets: exercise.sets,
      reps: exercise.reps,
      load: exercise.load,
      notes: "",
    });
  });

  return exercisesToInsert.length;
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
    bodyMapCategory: String(training?.bodyMapCategory ?? ""),
    bodyMapView: training?.bodyMapView === "back" ? "back" : "front",
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
  return EXERCISE_CATEGORY_BY_NAME.get(normalizeLookupValue(name)) || "";
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
