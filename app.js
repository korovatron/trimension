import * as THREE from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { Line2 } from 'three/addons/lines/Line2.js';
import { LineGeometry } from 'three/addons/lines/LineGeometry.js';
import { LineMaterial } from 'three/addons/lines/LineMaterial.js';

const titleScreen = document.getElementById('title-screen');
const mainApp = document.getElementById('main-app');
const startBtn = document.getElementById('start-btn');
const helpOverlay = document.getElementById('help-overlay');

let appInitialized = false;
let trimensionApp = null;

function setActualVH() {
    document.documentElement.style.setProperty('--actual-vh', `${window.innerHeight}px`);
}

function startApp() {
    titleScreen.classList.add('hidden');
    mainApp.style.display = 'block';

    if (!appInitialized) {
        trimensionApp = new TrimensionApp();
        window.trimensionApp = trimensionApp;
        appInitialized = true;
    }
}

function returnToTitleScreen() {
    titleScreen.classList.remove('hidden');
    mainApp.style.display = 'none';
    if (trimensionApp) {
        trimensionApp.cleanup();
        trimensionApp = null;
        window.trimensionApp = null;
    }
    appInitialized = false;
}

setActualVH();
window.addEventListener('resize', setActualVH);
window.addEventListener('orientationchange', () => window.setTimeout(setActualVH, 100));

startBtn.addEventListener('click', startApp);
document.getElementById('return-to-title').addEventListener('click', returnToTitleScreen);
document.getElementById('help-button').addEventListener('click', () => helpOverlay.classList.add('show'));
document.getElementById('close-help-btn').addEventListener('click', () => helpOverlay.classList.remove('show'));
helpOverlay.addEventListener('click', (event) => {
    if (event.target === helpOverlay) {
        helpOverlay.classList.remove('show');
    }
});

document.addEventListener('keydown', (event) => {
    if (event.code === 'Space' && !appInitialized) {
        event.preventDefault();
        startApp();
    } else if (event.code === 'Escape') {
        if (helpOverlay.classList.contains('show')) {
            helpOverlay.classList.remove('show');
            return;
        }

        if (appInitialized) {
            returnToTitleScreen();
        }
    }
});

const ATTACHMENT_FACES = {
    cuboid: [
        { id: 'top',    type: 'rectangle', normal: new THREE.Vector3(0, 1, 0),   center: (p) => new THREE.Vector3(0, p.height / 2, 0),   dims: ['width', 'depth'],  label: 'Top' },
        { id: 'bottom', type: 'rectangle', normal: new THREE.Vector3(0, -1, 0),  center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['width', 'depth'],  label: 'Bottom' },
        { id: 'front',  type: 'rectangle', normal: new THREE.Vector3(0, 0, 1),   center: (p) => new THREE.Vector3(0, 0, p.depth / 2),     dims: ['width', 'height'], label: 'Front' },
        { id: 'back',   type: 'rectangle', normal: new THREE.Vector3(0, 0, -1),  center: (p) => new THREE.Vector3(0, 0, -p.depth / 2),    dims: ['width', 'height'], label: 'Back' },
        { id: 'right',  type: 'rectangle', normal: new THREE.Vector3(1, 0, 0),   center: (p) => new THREE.Vector3(p.width / 2, 0, 0),     dims: ['depth', 'height'], label: 'Right' },
        { id: 'left',   type: 'rectangle', normal: new THREE.Vector3(-1, 0, 0),  center: (p) => new THREE.Vector3(-p.width / 2, 0, 0),    dims: ['depth', 'height'], label: 'Left' },
    ],
    'rectangular-pyramid': [
        { id: 'base',   type: 'rectangle', normal: new THREE.Vector3(0, -1, 0),  center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['length', 'width'], label: 'Base' },
    ],
    cylinder: [
        { id: 'top',    type: 'circle',    normal: new THREE.Vector3(0, 1, 0),   center: (p) => new THREE.Vector3(0, p.height / 2, 0),   dims: ['radius'],          label: 'Top' },
        { id: 'bottom', type: 'circle',    normal: new THREE.Vector3(0, -1, 0),  center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['radius'],          label: 'Bottom' },
    ],
    cone: [
        { id: 'base',   type: 'circle',    normal: new THREE.Vector3(0, -1, 0),  center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['radius'],          label: 'Base' },
    ],
    hemisphere: [
        { id: 'flat',   type: 'circle',    normal: new THREE.Vector3(0, -1, 0),  center: (p) => new THREE.Vector3(0, -p.radius / 2, 0),  dims: ['radius'],          label: 'Flat Face' },
    ],
    sphere: [],
    'right-triangle-prism': [],
    'trapezium-prism': [],
};

class TrimensionApp {
    constructor() {
        this.canvas = document.getElementById('three-canvas');
        this.panelToggleBtn = document.getElementById('panel-toggle-btn');
        this.controlPanel = document.querySelector('.control-panel');
        this.pointsListEl = document.getElementById('points-list');
        this.selectionSummaryEl = document.getElementById('selection-summary');
        this.actionsListEl = document.getElementById('actions-list');
        this.objectSections = {
            triangles: { header: document.getElementById('triangles-section-header'), content: document.getElementById('triangles-section-content'), arrow: document.getElementById('triangles-section-arrow'), list: document.getElementById('triangles-list'), count: document.getElementById('triangles-count') },
            segments:  { header: document.getElementById('segments-section-header'),  content: document.getElementById('segments-section-content'),  arrow: document.getElementById('segments-section-arrow'),  list: document.getElementById('segments-list'),  count: document.getElementById('segments-count')  },
            angles:    { header: document.getElementById('angles-section-header'),    content: document.getElementById('angles-section-content'),    arrow: document.getElementById('angles-section-arrow'),    list: document.getElementById('angles-list'),    count: document.getElementById('angles-count')    },
            planes:    { header: document.getElementById('planes-section-header'),    content: document.getElementById('planes-section-content'),    arrow: document.getElementById('planes-section-arrow'),    list: document.getElementById('planes-list'),    count: document.getElementById('planes-count')    },
            labels:    { header: document.getElementById('labels-section-header'),    content: document.getElementById('labels-section-content'),    arrow: document.getElementById('labels-section-arrow'),    list: document.getElementById('labels-list'),    count: document.getElementById('labels-count')    },
        };
        this.primitiveSelect = document.getElementById('primitive-select');
        this.primitiveCardsListEl = document.getElementById('primitive-cards-list');
        this.primitiveChip = document.getElementById('primitive-chip');
        this.orientationChip = document.getElementById('orientation-chip');
        this.ghostToggleBtn = document.getElementById('ghost-toggle-btn');
        this.addBtn = document.getElementById('add-btn');
        this.addDropdown = document.getElementById('add-dropdown');
        this.primitiveSectionHeader = document.getElementById('primitive-section-header');
        this.primitiveSectionContent = document.getElementById('primitive-section-content');
        this.primitiveSectionArrow = document.getElementById('primitive-section-arrow');
        this.pointsSectionHeader = document.getElementById('points-section-header');
        this.pointsSectionContent = document.getElementById('points-section-content');
        this.pointsSectionArrow = document.getElementById('points-section-arrow');

        this.panelOpen = true;
        this.ghostFaces = true;
        this.nextObjectId = 1;
        this.sceneObjects = [];
        this.selectedPoints = [];
        this.pointDefinitions = [];
        this.derivedPoints = [];
        this.pointMarkers = new Map();
        this.pointSprites = [];
        this.labelSprites = [];
        this.constructionLineMaterials = new Set();
        this.constructionPalette = [
            0xff595e,
            0xff924c,
            0xffca3a,
            0x8ac926,
            0x1982c4,
            0x6a4c93,
            0x2ec4b6,
            0xe76f51
        ];
        this.constructionColorIndex = 0;
        this.objectGroupCollapsed = {
            triangles: true,
            segments: true,
            angles: true,
            planes: true,
            labels: true
        };
        this.primitiveSectionCollapsed = true;
        this.pointsSectionCollapsed = true;

        this.defaultParams = {
            cuboid: { width: 7, depth: 4, height: 5 },
            'right-triangle-prism': { legA: 5, legB: 4, length: 7 },
            'trapezium-prism': { baseWidth: 6, leftHeight: 4, rightHeight: 2.5, length: 7 },
            sphere: { radius: 3 },
            hemisphere: { radius: 3 },
            cylinder: { radius: 2.5, height: 6 },
            cone: { radius: 2.5, height: 6 },
            'rectangular-pyramid': { length: 6.5, width: 4.5, height: 6 }
        };

        // compositeSlots: array of { id, primitive, orientation, params, hostFaceListIdx }
        this.compositeSlots = [
            { id: 0, primitive: 'cuboid', orientation: 'standard', params: { width: 7, depth: 4, height: 5 }, hostFaceListIdx: 0 }
        ];
        this.nextSlotId = 1;
        this.compositeGroup = null;
        this.slotGroupMap = new Map();   // slotId → Three.js Group
        this.primitiveMeshes = [];       // one mesh per slot
        this.slotLinkages = [];          // { fromSlotId, fromParam, toSlotId, toParam }

        this.orientations = {
            cuboid: [
                { value: 'standard', label: 'Standard' }
            ],
            'right-triangle-prism': [
                { value: 'standard', label: 'Standard' }
            ],
            'trapezium-prism': [
                { value: 'standard', label: 'Standard' }
            ],
            sphere: [
                { value: 'standard', label: 'Standard' }
            ],
            hemisphere: [
                { value: 'standard', label: 'Standard' }
            ],
            cylinder: [
                { value: 'vertical', label: 'Vertical', chipLabel: 'Vert' },
                { value: 'horizontal', label: 'Horizontal', chipLabel: 'Horiz' }
            ],
            cone: [
                { value: 'apex-up', label: 'Apex Up', chipLabel: 'Apex Up' },
                { value: 'apex-down', label: 'Apex Down', chipLabel: 'Apex Dn' },
                { value: 'sideways-right', label: 'Sideways Right', chipLabel: 'Side' }
            ],
            'rectangular-pyramid': [
                { value: 'apex-up', label: 'Apex Up', chipLabel: 'Apex Up' },
                { value: 'apex-down', label: 'Apex Down', chipLabel: 'Apex Dn' }
            ]
        };

        this.primitiveMeta = {
            cuboid: {
                label: 'Cuboid',
                params: [
                    { key: 'width', label: 'Length', min: 2, max: 10, step: 0.5 },
                    { key: 'depth', label: 'Width', min: 2, max: 10, step: 0.5 },
                    { key: 'height', label: 'Height', min: 2, max: 10, step: 0.5 }
                ]
            },
            'right-triangle-prism': {
                label: 'Right-Triangle Prism',
                params: [
                    { key: 'legA', label: 'Triangle Leg A', min: 2, max: 10, step: 0.5 },
                    { key: 'legB', label: 'Triangle Leg B', min: 2, max: 10, step: 0.5 },
                    { key: 'length', label: 'Prism Length', min: 2, max: 12, step: 0.5 }
                ]
            },
            'trapezium-prism': {
                label: 'Trapezium Prism',
                params: [
                    { key: 'baseWidth', label: 'Base Width', min: 2, max: 12, step: 0.5 },
                    { key: 'leftHeight', label: 'Left Height', min: 0, max: 10, step: 0.5 },
                    { key: 'rightHeight', label: 'Right Height', min: 0, max: 10, step: 0.5 },
                    { key: 'length', label: 'Prism Length', min: 2, max: 12, step: 0.5 }
                ]
            },
            sphere: {
                label: 'Sphere',
                params: [
                    { key: 'radius', label: 'Radius', min: 1, max: 6, step: 0.25 }
                ]
            },
            hemisphere: {
                label: 'Hemisphere',
                params: [
                    { key: 'radius', label: 'Radius', min: 1, max: 6, step: 0.25 }
                ]
            },
            cylinder: {
                label: 'Cylinder',
                params: [
                    { key: 'radius', label: 'Radius', min: 1, max: 5, step: 0.25 },
                    { key: 'height', label: 'Height', min: 2, max: 10, step: 0.5 }
                ]
            },
            cone: {
                label: 'Cone',
                params: [
                    { key: 'radius', label: 'Radius', min: 1, max: 5, step: 0.25 },
                    { key: 'height', label: 'Height', min: 2, max: 10, step: 0.5 }
                ]
            },
            'rectangular-pyramid': {
                label: 'Rectangular Pyramid',
                params: [
                    { key: 'length', label: 'Length', min: 2, max: 12, step: 0.5 },
                    { key: 'width', label: 'Width', min: 2, max: 12, step: 0.5 },
                    { key: 'height', label: 'Height', min: 2, max: 10, step: 0.5 }
                ]
            }
        };

        this.actionsByCount = {
            1: [
                { key: 'point-label', label: 'Add Point Label' }
            ],
            2: [
                { key: 'segment', label: 'Add Segment' },
                { key: 'length-label', label: 'Add Length Label' }
            ],
            3: [
                { key: 'triangle', label: 'Add Triangle' },
                { key: 'angle', label: 'Add Angle' }
            ],
            4: [
                { key: 'plane', label: 'Add Plane' }
            ]
        };

        this.handleWindowResize = this.onWindowResize.bind(this);
        this.handleCanvasPointerDown = this.handleCanvasPointerDown.bind(this);

        // Ensure startup collapse state and arrows are consistent
        this.primitiveSectionContent.classList.toggle('collapsed', this.primitiveSectionCollapsed);
        this.primitiveSectionHeader.setAttribute('aria-expanded', this.primitiveSectionCollapsed ? 'false' : 'true');
        this.primitiveSectionArrow.textContent = this.primitiveSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';

        this.pointsSectionContent.classList.toggle('collapsed', this.pointsSectionCollapsed);
        this.pointsSectionHeader.setAttribute('aria-expanded', this.pointsSectionCollapsed ? 'false' : 'true');
        this.pointsSectionArrow.textContent = this.pointsSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';

        Object.entries(this.objectSections).forEach(([key, sec]) => {
            const collapsed = this.objectGroupCollapsed[key];
            sec.content.classList.toggle('collapsed', collapsed);
            sec.header.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
            sec.arrow.textContent = collapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
        });

        this.initThree();
        this.bindEvents();
        this.buildComposite();
        this.renderCompositeCards();
        this.animate();
    }

    initThree() {
        this.scene = new THREE.Scene();
        this.scene.background = new THREE.Color(0xffffff);

        this.camera = new THREE.PerspectiveCamera(55, this.canvas.clientWidth / this.canvas.clientHeight, 0.1, 200);
        this.camera.position.set(10, 8, 11);

        this.renderer = new THREE.WebGLRenderer({ canvas: this.canvas, antialias: true, alpha: false });
        this.renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
        this.renderer.setSize(this.canvas.clientWidth, this.canvas.clientHeight, false);

        this.controls = new OrbitControls(this.camera, this.canvas);
        this.controls.enableDamping = true;
        this.controls.dampingFactor = 0.08;
        this.controls.target.set(0, 0, 0);
        this.controls.minDistance = 4;
        this.controls.maxDistance = 40;
        this.controls.touches = {
            ONE: THREE.TOUCH.ROTATE,
            TWO: THREE.TOUCH.DOLLY_PAN
        };

        const ambient = new THREE.AmbientLight(0xffffff, 1.1);
        const keyLight = new THREE.DirectionalLight(0xffffff, 1.35);
        keyLight.position.set(8, 14, 10);
        const fillLight = new THREE.DirectionalLight(0x9cc4ff, 0.5);
        fillLight.position.set(-8, 4, -6);
        this.scene.add(ambient, keyLight, fillLight);

        this.grid = new THREE.GridHelper(26, 26, 0xd9d9d9, 0xebebeb);
        this.grid.position.y = -4.5;
        this.scene.add(this.grid);

        window.addEventListener('resize', this.handleWindowResize);
    }

    bindEvents() {
        this.panelToggleBtn.addEventListener('click', () => {
            this.panelOpen = !this.panelOpen;
            this.controlPanel.classList.toggle('closed', !this.panelOpen);
            this.panelToggleBtn.classList.toggle('active', this.panelOpen);
        });

        this.canvas.addEventListener('pointerdown', this.handleCanvasPointerDown, { passive: true });

        this.addBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            // Populate dropdown dynamically based on current composite state
            const isFirst = this.compositeSlots.length === 0;
            const compatible = isFirst
                ? Object.keys(this.defaultParams)
                : this.getCompatiblePrimitives();

            this.addDropdown.innerHTML = '';

            if (!isFirst && (this.compositeSlots.length >= 3 || compatible.length === 0)) {
                const msg = document.createElement('div');
                msg.className = 'dropdown-item dropdown-item-disabled';
                msg.textContent = this.compositeSlots.length >= 3 ? 'Maximum 3 primitives' : 'No compatible additions';
                this.addDropdown.appendChild(msg);
            } else {
                compatible.forEach((primKey) => {
                    const item = document.createElement('div');
                    item.className = 'dropdown-item';
                    item.dataset.primitive = primKey;
                    item.innerHTML = `<strong>+</strong> Add ${this.primitiveMeta[primKey].label}`;
                    this.addDropdown.appendChild(item);
                });
            }

            const isOpening = this.addDropdown.style.display === 'none';
            this.addDropdown.style.display = isOpening ? 'block' : 'none';
        });

        this.addDropdown.addEventListener('click', (event) => {
            event.stopPropagation();
            const item = event.target.closest('[data-primitive]');
            if (!item) return;
            this.addSlot(item.dataset.primitive);
            this.addDropdown.style.display = 'none';
        });

        document.addEventListener('click', (event) => {
            if (event.target.closest('#add-btn') || event.target.closest('#add-dropdown')) return;
            this.addDropdown.style.display = 'none';
        });

        // Card-level delegation: orientation chips, cycle face, remove slot
        this.primitiveCardsListEl.addEventListener('click', (event) => {
            const chip = event.target.closest('[data-orientation-value]');
            if (chip) {
                const card = chip.closest('[data-slot-id]');
                if (!card) return;
                const slot = this.compositeSlots.find((s) => s.id === Number(card.dataset.slotId));
                if (!slot || chip.dataset.orientationValue === slot.orientation) return;
                slot.orientation = chip.dataset.orientationValue;
                card.querySelectorAll('[data-orientation-value]').forEach((btn) => {
                    btn.classList.toggle('is-active', btn.dataset.orientationValue === slot.orientation);
                    btn.setAttribute('aria-checked', btn.dataset.orientationValue === slot.orientation ? 'true' : 'false');
                });
                this.resetSceneObjects();
                this.buildComposite();
                return;
            }

            const cycleBtn = event.target.closest('[data-cycle-slot-id]');
            if (cycleBtn) {
                this.cycleSlotFace(Number(cycleBtn.dataset.cycleSlotId));
                return;
            }

            const removeBtn = event.target.closest('[data-remove-slot-id]');
            if (removeBtn) {
                this.removeSlot(Number(removeBtn.dataset.removeSlotId));
                return;
            }
        });

        this.pointsListEl.addEventListener('click', (event) => {
            const button = event.target.closest('[data-point-id]');
            if (!button) return;

            this.togglePointSelection(button.dataset.pointId);
        });

        this.actionsListEl.addEventListener('click', (event) => {
            const button = event.target.closest('[data-action-key]');
            if (!button) return;

            this.runAction(button.dataset.actionKey);
        });

        ['triangles', 'segments', 'angles', 'planes', 'labels'].forEach((key) => {
            const sec = this.objectSections[key];
            sec.list.addEventListener('click', (event) => {
                const toggleButton = event.target.closest('[data-toggle-object-id]');
                const deleteButton = event.target.closest('[data-delete-object-id]');
                if (toggleButton) this.toggleObjectVisibility(Number(toggleButton.dataset.toggleObjectId));
                if (deleteButton) this.deleteObject(Number(deleteButton.dataset.deleteObjectId));
            });
            const toggleSection = () => {
                this.objectGroupCollapsed[key] = !this.objectGroupCollapsed[key];
                const collapsed = this.objectGroupCollapsed[key];
                sec.content.classList.toggle('collapsed', collapsed);
                sec.header.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                sec.arrow.textContent = collapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
            };
            sec.header.addEventListener('click', toggleSection);
            sec.header.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection(); }
            });
        });

        document.getElementById('clear-selection-btn').addEventListener('click', () => this.clearSelection());
        document.getElementById('clear-objects-btn').addEventListener('click', () => this.clearObjects());
        document.getElementById('reset-view-btn').addEventListener('click', () => {
            this.resetView();
            this.closePanelOnMobile();
        });
        this.ghostToggleBtn.addEventListener('click', () => {
            this.ghostFaces = !this.ghostFaces;
            this.updatePrimitiveMaterial();
            this.ghostToggleBtn.textContent = this.ghostFaces ? 'GHOST\nFACES' : 'SOLID\nFACES';
        });

        const togglePrimitiveSection = () => {
            this.primitiveSectionCollapsed = !this.primitiveSectionCollapsed;
            this.primitiveSectionContent.classList.toggle('collapsed', this.primitiveSectionCollapsed);
            this.primitiveSectionHeader.setAttribute('aria-expanded', this.primitiveSectionCollapsed ? 'false' : 'true');
            this.primitiveSectionArrow.textContent = this.primitiveSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
        };
        this.primitiveSectionHeader.addEventListener('click', togglePrimitiveSection);
        this.primitiveSectionHeader.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); togglePrimitiveSection(); }
        });

        const togglePointsSection = () => {
            this.pointsSectionCollapsed = !this.pointsSectionCollapsed;
            this.pointsSectionContent.classList.toggle('collapsed', this.pointsSectionCollapsed);
            this.pointsSectionHeader.setAttribute('aria-expanded', this.pointsSectionCollapsed ? 'false' : 'true');
            this.pointsSectionArrow.textContent = this.pointsSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
        };

        this.pointsSectionHeader.addEventListener('click', (event) => {
            if (event.target.closest('#clear-selection-btn')) {
                return;
            }
            togglePointsSection();
        });

        this.pointsSectionHeader.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                togglePointsSection();
            }
        });
    }

    handleCanvasPointerDown(event) {
        if (event.target.closest('.control-panel')) {
            return;
        }

        if (window.innerWidth < 768 && this.panelOpen) {
            this.closePanelOnMobile();
        }
    }

    closePanelOnMobile() {
        if (window.innerWidth >= 768 || !this.panelOpen) {
            return;
        }

        this.panelOpen = false;
        this.controlPanel.classList.add('closed');
        this.panelToggleBtn.classList.remove('active');
    }

    onWindowResize() {
        const width = this.canvas.clientWidth;
        const height = this.canvas.clientHeight;
        this.camera.aspect = width / height;
        this.camera.updateProjectionMatrix();
        this.renderer.setSize(width, height, false);
        this.updateConstructionLineMaterialResolutions();
        this.refreshSliderBadges();
    }

    updateConstructionLineMaterialResolutions() {
        const width = this.canvas.clientWidth;
        const height = this.canvas.clientHeight;
        this.constructionLineMaterials.forEach((material) => {
            material.resolution.set(width, height);
        });
    }

    // --- Composite helpers ---

    getOrientationQuaternion(primitiveKey, orientationValue) {
        const q = new THREE.Quaternion();
        
        if (primitiveKey === 'cylinder') {
            if (orientationValue === 'horizontal') {
                q.setFromAxisAngle(new THREE.Vector3(0, 0, 1), Math.PI / 2);
            }
        } else if (primitiveKey === 'cone') {
            if (orientationValue === 'apex-down') {
                q.setFromAxisAngle(new THREE.Vector3(0, 0, 1), Math.PI);
            } else if (orientationValue === 'sideways-right') {
                q.setFromAxisAngle(new THREE.Vector3(0, 0, 1), -Math.PI / 2);
            }
        } else if (primitiveKey === 'rectangular-pyramid') {
            if (orientationValue === 'apex-down') {
                q.setFromAxisAngle(new THREE.Vector3(0, 0, 1), Math.PI);
            }
        }
        
        return q;
    }

    getCompatiblePrimitives() {
        if (this.compositeSlots.length >= 3) return [];

        const existingFaceTypes = new Set();
        this.compositeSlots.forEach((slot) => {
            (ATTACHMENT_FACES[slot.primitive] || []).forEach((f) => existingFaceTypes.add(f.type));
        });

        const existingPrimitives = new Set(this.compositeSlots.map((s) => s.primitive));

        return Object.keys(this.defaultParams).filter((primKey) => {
            if (existingPrimitives.has(primKey)) return false;
            const guestFaces = ATTACHMENT_FACES[primKey] || [];
            if (guestFaces.length === 0) return false;
            return guestFaces.some((gf) => existingFaceTypes.has(gf.type));
        });
    }

    getValidHostFaceEntries(guestSlot, previousSlots) {
        const guestFaceTypes = new Set((ATTACHMENT_FACES[guestSlot.primitive] || []).map((f) => f.type));
        if (guestFaceTypes.size === 0) return [];

        const result = [];
        previousSlots.forEach((hostSlot) => {
            (ATTACHMENT_FACES[hostSlot.primitive] || []).forEach((faceDef) => {
                if (guestFaceTypes.has(faceDef.type)) {
                    result.push({
                        slotId: hostSlot.id,
                        slot: hostSlot,
                        faceId: faceDef.id,
                        faceDef,
                        label: `${this.primitiveMeta[hostSlot.primitive].label}: ${faceDef.label}`,
                    });
                }
            });
        });
        return result;
    }

    getGuestAttachFaceDef(guestSlot, hostFaceNormal) {
        const faces = ATTACHMENT_FACES[guestSlot.primitive] || [];
        if (faces.length === 0) return null;

        const targetNormal = hostFaceNormal.clone().negate();
        let best = null;
        let bestDot = -Infinity;
        faces.forEach((face) => {
            const dot = face.normal.dot(targetNormal);
            if (dot > bestDot) { bestDot = dot; best = face; }
        });
        return best;
    }

    snapSlotDimensions(guestSlot) {
        const prevSlots = this.compositeSlots.filter((s) => s.id !== guestSlot.id);
        const entries = this.getValidHostFaceEntries(guestSlot, prevSlots);
        if (entries.length === 0) return;

        const hfIdx = guestSlot.hostFaceListIdx % entries.length;
        const { slot: hostSlot, faceDef: hostFaceDef } = entries[hfIdx];
        const guestFaceDef = (ATTACHMENT_FACES[guestSlot.primitive] || []).find((f) => f.type === hostFaceDef.type);
        if (!guestFaceDef) return;

        hostFaceDef.dims.forEach((dim, i) => {
            const guestDim = guestFaceDef.dims[i];
            if (guestDim) guestSlot.params[guestDim] = hostSlot.params[dim];
        });
    }

    addSlot(primitiveKey) {
        if (!this.primitiveMeta[primitiveKey]) return;

        if (this.compositeSlots.length === 0) {
            // First slot — reset everything
            const slot = {
                id: this.nextSlotId++,
                primitive: primitiveKey,
                orientation: this.orientations[primitiveKey][0].value,
                params: { ...this.defaultParams[primitiveKey] },
                hostFaceListIdx: 0,
            };
            this.compositeSlots.push(slot);
        } else {
            if (this.compositeSlots.length >= 3) return;
            const slot = {
                id: this.nextSlotId++,
                primitive: primitiveKey,
                orientation: this.orientations[primitiveKey][0].value,
                params: { ...this.defaultParams[primitiveKey] },
                hostFaceListIdx: 0,
            };
            this.compositeSlots.push(slot);
            this.snapSlotDimensions(slot);
        }

        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
        this.closePanelOnMobile();
    }

    removeSlot(slotId) {
        const idx = this.compositeSlots.findIndex((s) => s.id === slotId);
        if (idx === -1) return;
        // Remove this slot and any that follow (dependent chain)
        this.compositeSlots.splice(idx);
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    cycleSlotFace(slotId) {
        const slotIdx = this.compositeSlots.findIndex((s) => s.id === slotId);
        if (slotIdx <= 0) return;
        const slot = this.compositeSlots[slotIdx];
        const prevSlots = this.compositeSlots.slice(0, slotIdx);
        const entries = this.getValidHostFaceEntries(slot, prevSlots);
        if (entries.length <= 1) return;

        slot.hostFaceListIdx = (slot.hostFaceListIdx + 1) % entries.length;
        this.snapSlotDimensions(slot);
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    renderCompositeCards() {
        this.primitiveCardsListEl.innerHTML = '';
        const hasMultiple = this.compositeSlots.length > 1;

        if (this.compositeSlots.length === 0) {
            const empty = document.createElement('p');
            empty.className = 'section-note';
            empty.textContent = 'Click Add to create a primitive.';
            this.primitiveCardsListEl.appendChild(empty);
            return;
        }

        this.compositeSlots.forEach((slot, idx) => {
            if (idx > 0) {
                const cardSep = document.createElement('div');
                cardSep.className = 'primitive-card-separator';
                this.primitiveCardsListEl.appendChild(cardSep);
            }

            const card = document.createElement('div');
            card.className = 'primitive-card';
            card.dataset.slotId = String(slot.id);

            // Header
            const header = document.createElement('div');
            header.className = 'primitive-card-header';

            const titleEl = document.createElement('span');
            titleEl.className = 'primitive-card-title';
            titleEl.textContent = this.primitiveMeta[slot.primitive].label;
            header.appendChild(titleEl);

            if (hasMultiple) {
                const removeBtn = document.createElement('button');
                removeBtn.type = 'button';
                removeBtn.className = 'card-remove-btn';
                removeBtn.dataset.removeSlotId = String(slot.id);
                removeBtn.setAttribute('aria-label', `Remove ${this.primitiveMeta[slot.primitive].label}`);
                removeBtn.textContent = '×';
                header.appendChild(removeBtn);
            }
            card.appendChild(header);

            // Orientation chips (only if multiple options)
            const orientOptions = this.orientations[slot.primitive] || [];
            if (orientOptions.length > 1) {
                const orientRow = document.createElement('div');
                orientRow.className = 'orientation-inline-row';
                const orientLabel = document.createElement('span');
                orientLabel.className = 'field-label orientation-inline-label';
                orientLabel.textContent = 'Orientation';

                const chipsEl = document.createElement('div');
                chipsEl.className = 'orientation-chip-row';
                chipsEl.setAttribute('role', 'radiogroup');
                chipsEl.setAttribute('aria-label', 'Orientation');

                orientOptions.forEach((opt) => {
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'orientation-chip-btn';
                    btn.dataset.orientationValue = opt.value;
                    btn.textContent = opt.chipLabel || opt.label;
                    btn.title = opt.label;
                    btn.setAttribute('role', 'radio');
                    const isActive = opt.value === slot.orientation;
                    btn.classList.toggle('is-active', isActive);
                    btn.setAttribute('aria-checked', isActive ? 'true' : 'false');
                    chipsEl.appendChild(btn);
                });

                orientRow.appendChild(orientLabel);
                orientRow.appendChild(chipsEl);
                card.appendChild(orientRow);
            }

            // Sliders
            const sliderStack = document.createElement('div');
            sliderStack.className = 'slider-stack';

            (this.primitiveMeta[slot.primitive].params || []).forEach((config) => {
                const row = document.createElement('div');
                row.className = 'slider-row';
                row.dataset.paramKey = config.key;

                const labelEl = document.createElement('div');
                labelEl.className = 'slider-name';
                labelEl.textContent = config.label;

                const inputWrap = document.createElement('div');
                inputWrap.className = 'slider-input-wrap';

                const input = document.createElement('input');
                input.className = 'slider-input';
                input.type = 'range';
                input.min = String(config.min);
                input.max = String(config.max);
                input.step = String(config.step);
                input.value = String(slot.params[config.key]);

                const valueBadge = document.createElement('div');
                valueBadge.className = 'slider-value-badge';
                const valueText = document.createElement('span');
                valueText.className = 'slider-value-text';
                valueText.textContent = input.value;
                valueBadge.appendChild(valueText);

                input.addEventListener('input', () => {
                    const newVal = Number(input.value);
                    slot.params[config.key] = newVal;
                    valueText.textContent = input.value;
                    this.positionSliderBadge(input, valueBadge);

                    // Propagate to linked params
                    this.slotLinkages.forEach((link) => {
                        if (link.fromSlotId !== slot.id || link.fromParam !== config.key) return;
                        const linkedSlot = this.compositeSlots.find((s) => s.id === link.toSlotId);
                        if (!linkedSlot) return;
                        linkedSlot.params[link.toParam] = newVal;
                        const linkedRow = this.primitiveCardsListEl.querySelector(`[data-slot-id="${link.toSlotId}"] [data-param-key="${link.toParam}"]`);
                        if (linkedRow) {
                            const li = linkedRow.querySelector('.slider-input');
                            const lb = linkedRow.querySelector('.slider-value-badge');
                            const lt = lb?.querySelector('.slider-value-text');
                            if (li) li.value = String(newVal);
                            if (lt) lt.textContent = String(newVal);
                            if (li && lb) this.positionSliderBadge(li, lb);
                        }
                    });

                    this.buildComposite({ fitCamera: false });
                });

                inputWrap.append(input, valueBadge);
                row.append(labelEl, inputWrap);
                sliderStack.appendChild(row);
                requestAnimationFrame(() => this.positionSliderBadge(input, valueBadge));
            });

            card.appendChild(sliderStack);

            // Cycle face button (slot 1+)
            if (idx > 0) {
                const prevSlots = this.compositeSlots.slice(0, idx);
                const entries = this.getValidHostFaceEntries(slot, prevSlots);
                if (entries.length > 0) {
                    const hfIdx = slot.hostFaceListIdx % entries.length;
                    const currentLabel = entries[hfIdx]?.label || 'Face';

                    if (entries.length > 1) {
                        const cycleBtn = document.createElement('button');
                        cycleBtn.type = 'button';
                        cycleBtn.className = 'card-cycle-btn';
                        cycleBtn.dataset.cycleSlotId = String(slot.id);
                        cycleBtn.textContent = `↻ ${currentLabel}`;
                        card.appendChild(cycleBtn);
                    } else {
                        const faceInfo = document.createElement('div');
                        faceInfo.className = 'card-face-info';
                        faceInfo.textContent = currentLabel;
                        card.appendChild(faceInfo);
                    }
                }
            }

            this.primitiveCardsListEl.appendChild(card);
        });
    }

    positionSliderBadge(input, badge) {
        const min = Number(input.min);
        const max = Number(input.max);
        const value = Number(input.value);
        const ratio = max === min ? 0 : (value - min) / (max - min);

        input.style.setProperty('--slider-fill', `${Math.max(0, Math.min(1, ratio)) * 100}%`);

        const width = input.clientWidth || 1;
        const thumbSize = 18;
        const badgeHalf = (badge.offsetWidth || 24) / 2;
        const rawX = ratio * (width - thumbSize) + (thumbSize / 2);
        const x = Math.max(badgeHalf, Math.min(width - badgeHalf, rawX));
        badge.style.left = `${x}px`;
    }

    refreshSliderBadges() {
        this.primitiveCardsListEl.querySelectorAll('.slider-input-wrap').forEach((wrap) => {
            const input = wrap.querySelector('.slider-input');
            const badge = wrap.querySelector('.slider-value-badge');
            if (!input || !badge) return;
            this.positionSliderBadge(input, badge);
        });
    }

    buildPrimitive(options = {}) {
        this.buildComposite(options);
    }

    buildComposite(options = {}) {
        const fitCamera = options.fitCamera !== false;
        this.clearComposite();

        this.compositeGroup = new THREE.Group();
        this.scene.add(this.compositeGroup);
        this.primitiveGroup = this.compositeGroup; // alias used by buildPointMarkers etc.
        this.primitiveMeshes = [];
        this.slotLinkages = [];

        if (this.compositeSlots.length === 0) {
            this.pointDefinitions = [];
            this.rebuildConstructions();
            this.refreshDerivedPoints();
            this.buildPointMarkers();
            this.updatePanelCopy();
            this.renderPointsList();
            this.renderSelectionSummary();
            this.renderActions();
            return;
        }

        let allPoints = [];
        let maxBoundsRadius = 6;
        const usePrefix = this.compositeSlots.length > 1;

        this.compositeSlots.forEach((slot, idx) => {
            const def = this.createSlotDefinition(slot);

            if (idx > 0) {
                const prevSlots = this.compositeSlots.slice(0, idx);
                const entries = this.getValidHostFaceEntries(slot, prevSlots);
                if (entries.length > 0) {
                    const hfIdx = slot.hostFaceListIdx % entries.length;
                    const entry = entries[hfIdx];
                    const hostSlotGroup = this.slotGroupMap.get(entry.slotId);
                    const hostGroupQ = hostSlotGroup ? hostSlotGroup.quaternion : new THREE.Quaternion();
                    const hostGroupP = hostSlotGroup ? hostSlotGroup.position : new THREE.Vector3();

                    // Apply orientation rotation to face definition (which is in standard space)
                    const hostOrientQ = this.getOrientationQuaternion(entry.slot.primitive, entry.slot.orientation);

                    const hostFaceCenter = entry.faceDef.center(entry.slot.params)
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ)
                        .add(hostGroupP);
                    const hostFaceNormal = entry.faceDef.normal
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ);

                    this.applySlotTransform(def.group, slot, hostFaceCenter, hostFaceNormal);

                    const guestFaceDef = this.getGuestAttachFaceDef(slot, entry.faceDef.normal);
                    if (guestFaceDef) {
                        this.addLinkages(
                            { slotId: entry.slotId, dims: entry.faceDef.dims },
                            { slotId: slot.id, dims: guestFaceDef.dims }
                        );
                    }
                }
            }

            this.slotGroupMap.set(slot.id, def.group);
            this.compositeGroup.add(def.group);
            this.primitiveMeshes.push(def.mesh);

            const qRot = def.group.quaternion;
            const vPos = def.group.position;
            const worldPoints = def.points.map((pt) => ({
                ...pt,
                id: usePrefix ? `s${slot.id}_${pt.id}` : pt.id,
                label: usePrefix ? `${idx + 1}${pt.label}` : pt.label,
                position: pt.position.clone().applyQuaternion(qRot).add(vPos),
            }));
            allPoints = allPoints.concat(worldPoints);
            maxBoundsRadius = Math.max(maxBoundsRadius, def.boundsRadius + def.group.position.length());
        });

        this.pointDefinitions = allPoints;
        this.rebuildConstructions();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.updatePanelCopy();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
        if (fitCamera) {
            this.fitCameraToPrimitive(maxBoundsRadius);
        }
    }

    applySlotTransform(slotGroup, slot, hostFaceCenter, hostFaceNormal) {
        const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal);
        if (!guestFaceDef) return;

        const targetGuestNormal = hostFaceNormal.clone().negate();
        
        // Apply guest's orientation rotation to its face normal (which is in standard space)
        const guestOrientQ = this.getOrientationQuaternion(slot.primitive, slot.orientation);
        const guestFaceNormal = guestFaceDef.normal.clone().applyQuaternion(guestOrientQ);
        
        const Q = new THREE.Quaternion();
        const dot = guestFaceNormal.dot(targetGuestNormal);

        if (dot > 0.9999) {
            Q.identity();
        } else if (dot < -0.9999) {
            const perp = Math.abs(guestFaceNormal.x) < 0.9
                ? new THREE.Vector3(1, 0, 0).cross(guestFaceNormal).normalize()
                : new THREE.Vector3(0, 1, 0).cross(guestFaceNormal).normalize();
            Q.setFromAxisAngle(perp, Math.PI);
        } else {
            Q.setFromUnitVectors(guestFaceNormal, targetGuestNormal);
        }

        slotGroup.quaternion.copy(Q);
        const guestLocalCenter = guestFaceDef.center(slot.params)
            .clone()
            .applyQuaternion(guestOrientQ);  // Apply orientation first
        const guestRotatedCenter = guestLocalCenter.applyQuaternion(Q);
        slotGroup.position.copy(hostFaceCenter).sub(guestRotatedCenter);
    }

    addLinkages(hostInfo, guestInfo) {
        const len = Math.min(hostInfo.dims.length, guestInfo.dims.length);
        for (let i = 0; i < len; i++) {
            if (hostInfo.dims[i] && guestInfo.dims[i]) {
                this.slotLinkages.push({ fromSlotId: hostInfo.slotId, fromParam: hostInfo.dims[i], toSlotId: guestInfo.slotId, toParam: guestInfo.dims[i] });
                this.slotLinkages.push({ fromSlotId: guestInfo.slotId, fromParam: guestInfo.dims[i], toSlotId: hostInfo.slotId, toParam: hostInfo.dims[i] });
            }
        }
    }

    updatePanelCopy() {
        if (this.compositeSlots.length === 0) {
            this.primitiveChip.textContent = '—';
            this.orientationChip.textContent = '—';
            return;
        }
        const slot0 = this.compositeSlots[0];
        const extra = this.compositeSlots.length - 1;
        this.primitiveChip.textContent = extra > 0
            ? `${this.primitiveMeta[slot0.primitive].label} +${extra}`
            : this.primitiveMeta[slot0.primitive].label;
        const orientationLabel = this.orientations[slot0.primitive].find((o) => o.value === slot0.orientation)?.label || 'Standard';
        this.orientationChip.textContent = extra > 0 ? 'Composite' : orientationLabel;
    }

    createSlotDefinition(slot) {
        const group = new THREE.Group();
        const primitiveKey = slot.primitive;
        const params = slot.params;
        let points = [];
        let geometry;
        let boundsRadius = 6;
        let guideCircles = [];

        if (primitiveKey === 'cuboid') {
            const { width, depth, height } = params;
            geometry = new THREE.BoxGeometry(width, height, depth);
            points = [
                { id: 'A', label: 'A', description: 'bottom front left', position: new THREE.Vector3(-width / 2, -height / 2, depth / 2) },
                { id: 'B', label: 'B', description: 'bottom front right', position: new THREE.Vector3(width / 2, -height / 2, depth / 2) },
                { id: 'C', label: 'C', description: 'bottom back right', position: new THREE.Vector3(width / 2, -height / 2, -depth / 2) },
                { id: 'D', label: 'D', description: 'bottom back left', position: new THREE.Vector3(-width / 2, -height / 2, -depth / 2) },
                { id: 'E', label: 'E', description: 'top front left', position: new THREE.Vector3(-width / 2, height / 2, depth / 2) },
                { id: 'F', label: 'F', description: 'top front right', position: new THREE.Vector3(width / 2, height / 2, depth / 2) },
                { id: 'G', label: 'G', description: 'top back right', position: new THREE.Vector3(width / 2, height / 2, -depth / 2) },
                { id: 'H', label: 'H', description: 'top back left', position: new THREE.Vector3(-width / 2, height / 2, -depth / 2) }
            ];
            boundsRadius = Math.max(width, depth, height) * 1.15;
        } else if (primitiveKey === 'right-triangle-prism') {
            const { legA, legB, length } = params;
            const zFront = length / 2;
            const zBack = -length / 2;
            const xMin = -legA / 2;
            const xMax = legA / 2;
            const yMin = -legB / 2;
            const yMax = legB / 2;

            const vertices = [
                // Front triangle: A B C
                xMin, yMin, zFront, // A
                xMax, yMin, zFront, // B
                xMin, yMax, zFront, // C
                // Back triangle: D E F
                xMin, yMin, zBack,  // D
                xMax, yMin, zBack,  // E
                xMin, yMax, zBack   // F
            ];

            const indices = [
                // Front and back faces
                0, 1, 2,
                3, 5, 4,
                // Rectangle ABED
                0, 1, 4,
                0, 4, 3,
                // Rectangle ACFD
                0, 3, 5,
                0, 5, 2,
                // Rectangle BCEF
                1, 2, 5,
                1, 5, 4
            ];

            geometry = new THREE.BufferGeometry();
            geometry.setAttribute('position', new THREE.Float32BufferAttribute(vertices, 3));
            geometry.setIndex(indices);
            geometry.computeVertexNormals();

            points = [
                { id: 'A', label: 'A', description: 'front right-angle vertex', position: new THREE.Vector3(xMin, yMin, zFront) },
                { id: 'B', label: 'B', description: 'front base vertex', position: new THREE.Vector3(xMax, yMin, zFront) },
                { id: 'C', label: 'C', description: 'front height vertex', position: new THREE.Vector3(xMin, yMax, zFront) },
                { id: 'D', label: 'D', description: 'back right-angle vertex', position: new THREE.Vector3(xMin, yMin, zBack) },
                { id: 'E', label: 'E', description: 'back base vertex', position: new THREE.Vector3(xMax, yMin, zBack) },
                { id: 'F', label: 'F', description: 'back height vertex', position: new THREE.Vector3(xMin, yMax, zBack) }
            ];

            boundsRadius = Math.max(legA, legB, length) * 1.2;
        } else if (primitiveKey === 'trapezium-prism') {
            const { baseWidth, leftHeight, rightHeight, length } = params;
            const zFront = length / 2;
            const zBack = -length / 2;
            const xLeft = -baseWidth / 2;
            const xRight = baseWidth / 2;
            const maxHeight = Math.max(leftHeight, rightHeight, 0.001);
            const yBase = -maxHeight / 2;

            const vertices = [
                // Front trapezium: A B C D
                xLeft, yBase, zFront,                  // A
                xRight, yBase, zFront,                 // B
                xRight, yBase + rightHeight, zFront,   // C
                xLeft, yBase + leftHeight, zFront,     // D
                // Back trapezium: E F G H
                xLeft, yBase, zBack,                   // E
                xRight, yBase, zBack,                  // F
                xRight, yBase + rightHeight, zBack,    // G
                xLeft, yBase + leftHeight, zBack       // H
            ];

            const indices = [
                // Front and back faces
                0, 1, 2,
                0, 2, 3,
                4, 7, 6,
                4, 6, 5,
                // Side faces
                0, 1, 5,
                0, 5, 4,
                1, 2, 6,
                1, 6, 5,
                2, 3, 7,
                2, 7, 6,
                3, 0, 4,
                3, 4, 7
            ];

            geometry = new THREE.BufferGeometry();
            geometry.setAttribute('position', new THREE.Float32BufferAttribute(vertices, 3));
            geometry.setIndex(indices);
            geometry.computeVertexNormals();

            points = [
                { id: 'A', label: 'A', description: 'front bottom left', position: new THREE.Vector3(xLeft, yBase, zFront) },
                { id: 'B', label: 'B', description: 'front bottom right', position: new THREE.Vector3(xRight, yBase, zFront) },
                { id: 'C', label: 'C', description: 'front top right', position: new THREE.Vector3(xRight, yBase + rightHeight, zFront) },
                { id: 'D', label: 'D', description: 'front top left', position: new THREE.Vector3(xLeft, yBase + leftHeight, zFront) },
                { id: 'E', label: 'E', description: 'back bottom left', position: new THREE.Vector3(xLeft, yBase, zBack) },
                { id: 'F', label: 'F', description: 'back bottom right', position: new THREE.Vector3(xRight, yBase, zBack) },
                { id: 'G', label: 'G', description: 'back top right', position: new THREE.Vector3(xRight, yBase + rightHeight, zBack) },
                { id: 'H', label: 'H', description: 'back top left', position: new THREE.Vector3(xLeft, yBase + leftHeight, zBack) }
            ];

            const heightDelta = leftHeight - rightHeight;
            const shorterHeight = Math.min(leftHeight, rightHeight);
            if (Math.abs(heightDelta) > 1e-6 && shorterHeight > 1e-6) {
                const isLeftTaller = heightDelta > 0;
                const helperX = isLeftTaller ? xLeft : xRight;
                const helperY = yBase + shorterHeight;
                const helperSideDescription = isLeftTaller ? 'left' : 'right';

                points.push(
                    {
                        id: 'I',
                        label: 'I',
                        description: `front ${helperSideDescription} helper point`,
                        position: new THREE.Vector3(helperX, helperY, zFront)
                    },
                    {
                        id: 'J',
                        label: 'J',
                        description: `back ${helperSideDescription} helper point`,
                        position: new THREE.Vector3(helperX, helperY, zBack)
                    }
                );
            }

            boundsRadius = Math.max(baseWidth, leftHeight, rightHeight, length) * 1.2;
        } else if (primitiveKey === 'sphere') {
            const { radius } = params;
            geometry = new THREE.SphereGeometry(radius, 48, 32);
            points = [
                { id: 'A', label: 'A', description: 'top', position: new THREE.Vector3(0, radius, 0) },
                { id: 'B', label: 'B', description: 'bottom', position: new THREE.Vector3(0, -radius, 0) },
                { id: 'C', label: 'C', description: 'front', position: new THREE.Vector3(0, 0, radius) },
                { id: 'D', label: 'D', description: 'right', position: new THREE.Vector3(radius, 0, 0) },
                { id: 'E', label: 'E', description: 'back', position: new THREE.Vector3(0, 0, -radius) },
                { id: 'F', label: 'F', description: 'left', position: new THREE.Vector3(-radius, 0, 0) },
                { id: 'O', label: 'O', description: 'centre', position: new THREE.Vector3(0, 0, 0) }
            ];

            const segments = 128;
            const equator = [];
            const meridianYZ = [];
            const meridianXY = [];
            for (let i = 0; i < segments; i += 1) {
                const t = (i / segments) * Math.PI * 2;
                equator.push(new THREE.Vector3(radius * Math.cos(t), 0, radius * Math.sin(t)));
                meridianYZ.push(new THREE.Vector3(0, radius * Math.sin(t), radius * Math.cos(t)));
                meridianXY.push(new THREE.Vector3(radius * Math.cos(t), radius * Math.sin(t), 0));
            }
            guideCircles = [equator, meridianYZ, meridianXY];

            boundsRadius = radius * 1.35;
        } else if (primitiveKey === 'hemisphere') {
            const { radius } = params;
            geometry = new THREE.SphereGeometry(radius, 48, 24, 0, Math.PI * 2, 0, Math.PI / 2);
            geometry.translate(0, -radius / 2, 0);
            points = [
                { id: 'A', label: 'A', description: 'dome top', position: new THREE.Vector3(0, radius / 2, 0) },
                { id: 'B', label: 'B', description: 'rim front', position: new THREE.Vector3(0, -radius / 2, radius) },
                { id: 'C', label: 'C', description: 'rim right', position: new THREE.Vector3(radius, -radius / 2, 0) },
                { id: 'D', label: 'D', description: 'rim back', position: new THREE.Vector3(0, -radius / 2, -radius) },
                { id: 'E', label: 'E', description: 'rim left', position: new THREE.Vector3(-radius, -radius / 2, 0) },
                { id: 'F', label: 'F', description: 'flat face centre', position: new THREE.Vector3(0, -radius / 2, 0) }
            ];
            boundsRadius = radius * 1.35;
        } else if (primitiveKey === 'cylinder') {
            const { radius, height } = params;
            geometry = new THREE.CylinderGeometry(radius, radius, height, 48, 1, false);
            if (slot.orientation === 'horizontal') {
                geometry.rotateZ(Math.PI / 2);
                points = [
                    { id: 'A', label: 'A', description: 'right centre', position: new THREE.Vector3(height / 2, 0, 0) },
                    { id: 'B', label: 'B', description: 'left centre', position: new THREE.Vector3(-height / 2, 0, 0) },
                    { id: 'C', label: 'C', description: 'right top', position: new THREE.Vector3(height / 2, radius, 0) },
                    { id: 'D', label: 'D', description: 'right bottom', position: new THREE.Vector3(height / 2, -radius, 0) },
                    { id: 'E', label: 'E', description: 'left top', position: new THREE.Vector3(-height / 2, radius, 0) },
                    { id: 'F', label: 'F', description: 'left bottom', position: new THREE.Vector3(-height / 2, -radius, 0) },
                    { id: 'O', label: 'O', description: 'midpoint', position: new THREE.Vector3(0, 0, 0) }
                ];
            } else {
                points = [
                    { id: 'A', label: 'A', description: 'top centre', position: new THREE.Vector3(0, height / 2, 0) },
                    { id: 'B', label: 'B', description: 'bottom centre', position: new THREE.Vector3(0, -height / 2, 0) },
                    { id: 'C', label: 'C', description: 'top front', position: new THREE.Vector3(0, height / 2, radius) },
                    { id: 'D', label: 'D', description: 'top back', position: new THREE.Vector3(0, height / 2, -radius) },
                    { id: 'E', label: 'E', description: 'bottom front', position: new THREE.Vector3(0, -height / 2, radius) },
                    { id: 'F', label: 'F', description: 'bottom back', position: new THREE.Vector3(0, -height / 2, -radius) },
                    { id: 'O', label: 'O', description: 'midpoint', position: new THREE.Vector3(0, 0, 0) }
                ];
            }
            boundsRadius = Math.max(radius * 2, height) * 1.15;
        } else if (primitiveKey === 'cone') {
            const { radius, height } = params;
            geometry = new THREE.ConeGeometry(radius, height, 48, 1, false);
            if (slot.orientation === 'apex-down') {
                geometry.rotateZ(Math.PI);
                points = [
                    { id: 'A', label: 'A', description: 'apex', position: new THREE.Vector3(0, -height / 2, 0) },
                    { id: 'B', label: 'B', description: 'base centre', position: new THREE.Vector3(0, height / 2, 0) },
                    { id: 'C', label: 'C', description: 'base front', position: new THREE.Vector3(0, height / 2, radius) },
                    { id: 'D', label: 'D', description: 'base right', position: new THREE.Vector3(radius, height / 2, 0) },
                    { id: 'E', label: 'E', description: 'base back', position: new THREE.Vector3(0, height / 2, -radius) },
                    { id: 'F', label: 'F', description: 'base left', position: new THREE.Vector3(-radius, height / 2, 0) }
                ];
            } else if (slot.orientation === 'sideways-right') {
                geometry.rotateZ(-Math.PI / 2);
                points = [
                    { id: 'A', label: 'A', description: 'apex', position: new THREE.Vector3(height / 2, 0, 0) },
                    { id: 'B', label: 'B', description: 'base centre', position: new THREE.Vector3(-height / 2, 0, 0) },
                    { id: 'C', label: 'C', description: 'base top', position: new THREE.Vector3(-height / 2, radius, 0) },
                    { id: 'D', label: 'D', description: 'base front', position: new THREE.Vector3(-height / 2, 0, radius) },
                    { id: 'E', label: 'E', description: 'base bottom', position: new THREE.Vector3(-height / 2, -radius, 0) },
                    { id: 'F', label: 'F', description: 'base back', position: new THREE.Vector3(-height / 2, 0, -radius) }
                ];
            } else {
                points = [
                    { id: 'A', label: 'A', description: 'apex', position: new THREE.Vector3(0, height / 2, 0) },
                    { id: 'B', label: 'B', description: 'base centre', position: new THREE.Vector3(0, -height / 2, 0) },
                    { id: 'C', label: 'C', description: 'base front', position: new THREE.Vector3(0, -height / 2, radius) },
                    { id: 'D', label: 'D', description: 'base right', position: new THREE.Vector3(radius, -height / 2, 0) },
                    { id: 'E', label: 'E', description: 'base back', position: new THREE.Vector3(0, -height / 2, -radius) },
                    { id: 'F', label: 'F', description: 'base left', position: new THREE.Vector3(-radius, -height / 2, 0) }
                ];
            }
            boundsRadius = Math.max(radius * 2, height) * 1.18;
        } else if (primitiveKey === 'rectangular-pyramid') {
            const { length, width, height } = params;
            geometry = new THREE.ConeGeometry(1, height, 4, 1, false);
            geometry.rotateY(Math.PI / 4);
            geometry.scale(length / Math.sqrt(2), 1, width / Math.sqrt(2));
            if (slot.orientation === 'apex-down') {
                geometry.rotateZ(Math.PI);
                points = [
                    { id: 'A', label: 'A', description: 'top front left', position: new THREE.Vector3(-length / 2, height / 2, width / 2) },
                    { id: 'B', label: 'B', description: 'top front right', position: new THREE.Vector3(length / 2, height / 2, width / 2) },
                    { id: 'C', label: 'C', description: 'top back right', position: new THREE.Vector3(length / 2, height / 2, -width / 2) },
                    { id: 'D', label: 'D', description: 'top back left', position: new THREE.Vector3(-length / 2, height / 2, -width / 2) },
                    { id: 'E', label: 'E', description: 'apex', position: new THREE.Vector3(0, -height / 2, 0) },
                    { id: 'O', label: 'O', description: 'base centre', position: new THREE.Vector3(0, height / 2, 0) }
                ];
            } else {
                points = [
                    { id: 'A', label: 'A', description: 'base front left', position: new THREE.Vector3(-length / 2, -height / 2, width / 2) },
                    { id: 'B', label: 'B', description: 'base front right', position: new THREE.Vector3(length / 2, -height / 2, width / 2) },
                    { id: 'C', label: 'C', description: 'base back right', position: new THREE.Vector3(length / 2, -height / 2, -width / 2) },
                    { id: 'D', label: 'D', description: 'base back left', position: new THREE.Vector3(-length / 2, -height / 2, -width / 2) },
                    { id: 'E', label: 'E', description: 'apex', position: new THREE.Vector3(0, height / 2, 0) },
                    { id: 'O', label: 'O', description: 'base centre', position: new THREE.Vector3(0, -height / 2, 0) }
                ];
            }
            boundsRadius = Math.max(length, width, height) * 1.15;
        } else {
            throw new Error(`Unknown primitive key: ${primitiveKey}`);
        }

        const material = new THREE.MeshPhongMaterial({
            color: 0x7db3e8,
            transparent: true,
            opacity: this.ghostFaces ? 0.14 : 0.78,
            side: THREE.DoubleSide,
            shininess: 90,
            specular: 0x315579,
            depthWrite: false
        });

        const mesh = new THREE.Mesh(geometry, material);
        mesh.renderOrder = 5;
        const edgeGeometry = (primitiveKey === 'cone' || primitiveKey === 'cylinder' || primitiveKey === 'sphere' || primitiveKey === 'hemisphere')
            ? new THREE.EdgesGeometry(geometry, 30)
            : new THREE.EdgesGeometry(geometry);
        const edges = new THREE.LineSegments(
            edgeGeometry,
            new THREE.LineBasicMaterial({ color: 0x000000, transparent: false, opacity: 1 })
        );
        edges.renderOrder = 6;
        group.add(mesh, edges);

        if (guideCircles.length > 0) {
            guideCircles.forEach((circlePoints) => {
                const circleGeometry = new THREE.BufferGeometry().setFromPoints(circlePoints);
                const circle = new THREE.LineLoop(
                    circleGeometry,
                    new THREE.LineBasicMaterial({ color: 0x000000, transparent: false, opacity: 1 })
                );
                circle.renderOrder = 7;
                group.add(circle);
            });
        }

        return { group, mesh, points, boundsRadius };
    }

    updatePrimitiveMaterial() {
        this.primitiveMeshes.forEach((mesh) => {
            mesh.material.opacity = this.ghostFaces ? 0.14 : 0.78;
            mesh.material.needsUpdate = true;
        });
    }

    buildPointMarkers() {
        this.pointMarkers.forEach((marker) => {
            if (marker.parent) {
                marker.parent.remove(marker);
            }
            this.disposeObject3D(marker);
        });
        this.pointMarkers.clear();

        this.pointSprites.forEach((sprite) => {
            if (sprite.parent) {
                sprite.parent.remove(sprite);
            }
            this.disposeObject3D(sprite);
        });
        this.pointSprites = [];

        this.pointMarkers.clear();

        const markerGeometry = new THREE.SphereGeometry(0.1, 18, 18);
        this.getAllPoints().forEach((point) => {
            const markerColor = point.isDerived ? 0x2e7d32 : 0x000000;
            const marker = new THREE.Mesh(markerGeometry, new THREE.MeshBasicMaterial({ color: markerColor }));
            marker.position.copy(point.position);
            this.primitiveGroup.add(marker);
            this.pointMarkers.set(point.id, marker);

            const labelBackground = point.isDerived ? '#d9f9d6' : '#ffd84d';
            const sprite = this.createTextSprite(point.label, {
                fontSize: 52,
                textColor: '#000000',
                background: labelBackground,
                borderColor: '#000000'
            });
            sprite.position.copy(point.position.clone().add(new THREE.Vector3(0.18, 0.22, 0.18)));
            this.primitiveGroup.add(sprite);
            this.pointSprites.push(sprite);
        });
    }

    renderPointsList() {
        this.pointsListEl.innerHTML = '';
        this.getAllPoints().forEach((point) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'point-btn';
            if (point.isDerived) {
                button.classList.add('is-derived');
            }
            button.dataset.pointId = point.id;
            if (this.selectedPoints.includes(point.id)) {
                button.classList.add('is-selected');
            }
            button.setAttribute('aria-label', point.isDerived ? `${point.label}, derived point` : `Point ${point.label}`);
            button.title = point.isDerived ? `${point.label} - derived point` : point.label;
            button.innerHTML = `<span class="point-name">${point.label}</span>`;
            this.pointsListEl.appendChild(button);
        });
        this.updatePointMarkerStyles();
    }

    updatePointMarkerStyles() {
        this.getAllPoints().forEach((point) => {
            const marker = this.pointMarkers.get(point.id);
            if (!marker) return;

            const isSelected = this.selectedPoints.includes(point.id);
            const baseColor = point.isDerived ? 0x2e7d32 : 0x000000;
            marker.material.color.set(isSelected ? 0x4a90e2 : baseColor);
            marker.scale.setScalar(isSelected ? 1.45 : 1);
        });
    }

    togglePointSelection(pointId) {
        const existingIndex = this.selectedPoints.indexOf(pointId);
        if (existingIndex >= 0) {
            this.selectedPoints.splice(existingIndex, 1);
        } else {
            if (this.selectedPoints.length >= 4) {
                return;
            }
            this.selectedPoints.push(pointId);
        }

        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    clearSelection() {
        this.selectedPoints = [];
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    renderSelectionSummary() {
        this.selectionSummaryEl.innerHTML = '';
        if (this.selectedPoints.length === 0) {
            this.selectionSummaryEl.classList.add('empty');
            return;
        }

        this.selectionSummaryEl.classList.remove('empty');
        this.selectedPoints.forEach((pointId, index) => {
            const pill = document.createElement('span');
            pill.className = 'selection-pill';
            pill.textContent = pointId;
            this.selectionSummaryEl.appendChild(pill);

            if (index < this.selectedPoints.length - 1) {
                const arrow = document.createElement('span');
                arrow.className = 'selection-arrow';
                arrow.textContent = '→';
                this.selectionSummaryEl.appendChild(arrow);
            }
        });
    }

    renderActions() {
        this.actionsListEl.innerHTML = '';
        const actions = this.getValidActionsForSelection();

        if (actions.length === 0) {
            return;
        }

        actions.forEach((action) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'action-btn';
            button.dataset.actionKey = action.key;
            button.textContent = action.label;
            this.actionsListEl.appendChild(button);
        });
    }

    getValidActionsForSelection() {
        const baseActions = this.actionsByCount[this.selectedPoints.length] || [];
        if (this.selectedPoints.length !== 4) {
            return baseActions;
        }

        return baseActions.filter((action) => {
            if (action.key !== 'plane') {
                return true;
            }
            return this.areSelectedPointsCoplanar(this.selectedPoints);
        });
    }

    areSelectedPointsCoplanar(pointIds) {
        const vectors = this.getVectorsByPointIds(pointIds);
        if (!vectors || vectors.length !== 4) {
            return false;
        }

        const [p0, p1, p2, p3] = vectors;
        const v1 = p1.clone().sub(p0);
        const v2 = p2.clone().sub(p0);
        const normal = new THREE.Vector3().crossVectors(v1, v2);
        const normalLength = normal.length();

        // Degenerate quadruples (collinear/duplicate picks) do not define a plane region.
        if (normalLength < 1e-6) {
            return false;
        }

        const scale = Math.max(v1.length(), v2.length(), p3.distanceTo(p0), 1);
        const distanceToPlane = Math.abs(p3.clone().sub(p0).dot(normal.clone().normalize()));
        return distanceToPlane <= scale * 1e-4;
    }

    orderCoplanarPointIds(pointIds) {
        const points = pointIds.map((pointId) => this.getPointById(pointId)).filter(Boolean);
        if (points.length !== 4) {
            return pointIds;
        }

        const centroid = new THREE.Vector3();
        points.forEach((point) => centroid.add(point.position));
        centroid.multiplyScalar(1 / points.length);

        const baseVector = points[0].position.clone().sub(centroid);
        let normal = new THREE.Vector3();
        for (let index = 1; index < points.length - 1; index += 1) {
            const v1 = points[index].position.clone().sub(points[0].position);
            const v2 = points[index + 1].position.clone().sub(points[0].position);
            normal = new THREE.Vector3().crossVectors(v1, v2);
            if (normal.lengthSq() > 1e-8) {
                normal.normalize();
                break;
            }
        }

        if (normal.lengthSq() <= 1e-8 || baseVector.lengthSq() <= 1e-8) {
            return pointIds;
        }

        const axisX = baseVector.normalize();
        const axisY = new THREE.Vector3().crossVectors(normal, axisX).normalize();

        const ordered = points
            .map((point) => {
                const offset = point.position.clone().sub(centroid);
                const x = offset.dot(axisX);
                const y = offset.dot(axisY);
                return {
                    id: point.id,
                    angle: Math.atan2(y, x)
                };
            })
            .sort((left, right) => left.angle - right.angle)
            .map((entry) => entry.id);

        const originalFirstIndex = ordered.indexOf(pointIds[0]);
        if (originalFirstIndex <= 0) {
            return ordered;
        }

        return ordered.slice(originalFirstIndex).concat(ordered.slice(0, originalFirstIndex));
    }

    getPointById(pointId) {
        return this.getAllPoints().find((point) => point.id === pointId);
    }

    getAllPoints() {
        return [...this.pointDefinitions, ...this.derivedPoints];
    }

    refreshDerivedPoints() {
        const basePoints = new Map();
        this.pointDefinitions.forEach((point) => {
            basePoints.set(point.id, point.position.clone());
        });

        const usedLabels = new Set(this.pointDefinitions.map((point) => point.label));
        const candidateLetters = 'IJKLMNOPQRSTUVWXYZ'.split('').filter((letter) => !usedLabels.has(letter));
        let fallbackIndex = 1;

        const existingPointNear = (point, threshold = 0.02) => {
            for (const [, existingPoint] of basePoints.entries()) {
                if (existingPoint.distanceTo(point) <= threshold) {
                    return true;
                }
            }
            return false;
        };

        const nextDerivedLabel = () => {
            if (candidateLetters.length > 0) {
                const label = candidateLetters.shift();
                usedLabels.add(label);
                return label;
            }
            const label = `P${fallbackIndex}`;
            fallbackIndex += 1;
            return label;
        };

        const segments = this.getConstructionSegments(basePoints);
        const derived = [];

        for (let i = 0; i < segments.length; i += 1) {
            for (let j = i + 1; j < segments.length; j += 1) {
                const first = segments[i];
                const second = segments[j];

                const sharesEndpoint = first.pointIds.some((pointId) => second.pointIds.includes(pointId));
                if (sharesEndpoint) {
                    continue;
                }

                const intersection = this.findSegmentIntersection(first.start, first.end, second.start, second.end);
                if (!intersection) {
                    continue;
                }

                if (existingPointNear(intersection)) {
                    continue;
                }

                const label = nextDerivedLabel();
                const id = `derived-${label}`;
                basePoints.set(id, intersection.clone());
                derived.push({
                    id,
                    label,
                    description: 'derived intersection',
                    position: intersection,
                    isDerived: true
                });
            }
        }

        this.derivedPoints = derived;
        const validPointIds = new Set(this.getAllPoints().map((point) => point.id));
        this.selectedPoints = this.selectedPoints.filter((pointId) => validPointIds.has(pointId));
    }

    getConstructionSegments(pointMap) {
        const segments = [];

        const resolvePoint = (pointId) => pointMap.get(pointId);
        const pushSegment = (pointIdA, pointIdB) => {
            const start = resolvePoint(pointIdA);
            const end = resolvePoint(pointIdB);
            if (!start || !end) {
                return;
            }
            segments.push({ pointIds: [pointIdA, pointIdB], start, end });
        };

        this.sceneObjects.forEach((entry) => {
            const definition = entry.definition;
            if (!definition) {
                return;
            }

            if (definition.kind === 'segment' && definition.pointIds?.length === 2) {
                pushSegment(definition.pointIds[0], definition.pointIds[1]);
                return;
            }

            if (definition.kind === 'triangle' && definition.pointIds?.length === 3) {
                const ids = definition.pointIds;
                pushSegment(ids[0], ids[1]);
                pushSegment(ids[1], ids[2]);
                pushSegment(ids[2], ids[0]);
                return;
            }

            if (definition.kind === 'plane' && definition.pointIds?.length === 4) {
                const ids = definition.pointIds;
                pushSegment(ids[0], ids[1]);
                pushSegment(ids[1], ids[2]);
                pushSegment(ids[2], ids[3]);
                pushSegment(ids[3], ids[0]);
            }
        });

        return segments;
    }

    findSegmentIntersection(p1, p2, q1, q2) {
        const u = p2.clone().sub(p1);
        const v = q2.clone().sub(q1);
        const w0 = p1.clone().sub(q1);
        const a = u.dot(u);
        const b = u.dot(v);
        const c = v.dot(v);
        const d = u.dot(w0);
        const e = v.dot(w0);
        const denom = (a * c) - (b * b);

        if (Math.abs(denom) < 1e-8) {
            return null;
        }

        const s = ((b * e) - (c * d)) / denom;
        const t = ((a * e) - (b * d)) / denom;
        if (s < -1e-4 || s > 1 + 1e-4 || t < -1e-4 || t > 1 + 1e-4) {
            return null;
        }

        const pointOnFirst = p1.clone().add(u.multiplyScalar(s));
        const pointOnSecond = q1.clone().add(v.multiplyScalar(t));
        if (pointOnFirst.distanceTo(pointOnSecond) > 0.02) {
            return null;
        }

        return pointOnFirst.clone().add(pointOnSecond).multiplyScalar(0.5);
    }

    getSelectedVectors() {
        return this.selectedPoints.map((pointId) => this.getPointById(pointId)?.position.clone());
    }

    nextConstructionColor() {
        const color = this.constructionPalette[this.constructionColorIndex % this.constructionPalette.length];
        this.constructionColorIndex += 1;
        return color;
    }

    runAction(actionKey) {
        const vectors = this.getSelectedVectors();
        if (vectors.some((value) => !value)) {
            return;
        }

        if (actionKey === 'point-label') {
            const pointId = this.selectedPoints[0];
            const label = window.prompt(`Label for point ${pointId}`, pointId);
            if (!label) return;
            const point = this.getPointById(pointId);
            const sprite = this.createTextSprite(label, {
                fontSize: 46,
                textColor: '#000000',
                background: '#ffd84d',
                borderColor: '#000000'
            });
            sprite.position.copy(point.position.clone().add(new THREE.Vector3(0.34, 0.46, 0.18)));
            this.addSceneObject({
                type: 'label',
                name: `Label ${pointId}`,
                subtitle: label,
                object3D: sprite,
                definition: {
                    kind: 'point-label',
                    pointId,
                    text: label
                }
            });
        }

        if (actionKey === 'segment') {
            const ids = [...this.selectedPoints];
            const color = this.nextConstructionColor();
            const segment = this.createSegment(vectors[0], vectors[1], color);
            this.addSceneObject({
                type: 'segment',
                name: `Segment ${ids[0]}${ids[1]}`,
                subtitle: 'Two-point construction',
                object3D: segment,
                definition: {
                    kind: 'segment',
                    pointIds: ids,
                    color
                }
            });
        }

        if (actionKey === 'length-label') {
            const ids = [...this.selectedPoints];
            const label = window.prompt(`Label for ${ids[0]}${ids[1]}`, '23 cm');
            if (!label) return;
            const midpoint = vectors[0].clone().lerp(vectors[1], 0.5).add(new THREE.Vector3(0.2, 0.2, 0.2));
            const sprite = this.createTextSprite(label, {
                fontSize: 46,
                textColor: '#000000',
                background: '#b9f18a',
                borderColor: '#000000'
            });
            sprite.position.copy(midpoint);
            this.addSceneObject({
                type: 'label',
                name: `Label ${ids[0]}${ids[1]}`,
                subtitle: label,
                object3D: sprite,
                definition: {
                    kind: 'length-label',
                    pointIds: ids,
                    text: label
                }
            });
        }

        if (actionKey === 'triangle') {
            const ids = [...this.selectedPoints];
            const color = this.nextConstructionColor();
            const triangle = this.createTriangle(vectors[0], vectors[1], vectors[2], color, 0.28);
            this.addSceneObject({
                type: 'triangle',
                name: `Triangle ${ids.join('')}`,
                subtitle: 'Three-point section',
                object3D: triangle,
                definition: {
                    kind: 'triangle',
                    pointIds: ids,
                    color,
                    opacity: 0.28
                }
            });
        }

        if (actionKey === 'angle') {
            const ids = [...this.selectedPoints];
            const color = this.nextConstructionColor();
            const angleGroup = this.createAngleMarker(vectors[0], vectors[1], vectors[2], ids.join(''), color);
            this.addSceneObject({
                type: 'angle',
                name: `Angle ${ids.join('')}`,
                subtitle: `Angle at ${ids[1]}`,
                object3D: angleGroup,
                definition: {
                    kind: 'angle',
                    pointIds: ids,
                    color
                }
            });
        }

        if (actionKey === 'plane') {
            const ids = [...this.selectedPoints];
            if (!this.areSelectedPointsCoplanar(ids)) {
                return;
            }
            const orderedIds = this.orderCoplanarPointIds(ids);
            const orderedVectors = this.getVectorsByPointIds(orderedIds);
            const color = this.nextConstructionColor();
            const plane = this.createQuad(orderedVectors, color, 0.2);
            this.addSceneObject({
                type: 'plane',
                name: `Plane ${orderedIds.join('')}`,
                subtitle: 'Four-point shading',
                object3D: plane,
                definition: {
                    kind: 'plane',
                    pointIds: orderedIds,
                    color,
                    opacity: 0.2
                }
            });
        }

        this.clearSelection();
        this.closePanelOnMobile();
    }

    createSegment(start, end, color) {
        return this.createThickPolyline([start, end], color, 5);
    }

    createThickPolyline(points, color, width = 5) {
        const geometry = new LineGeometry();
        geometry.setPositions(points.flatMap((point) => [point.x, point.y, point.z]));

        const material = new LineMaterial({
            color,
            linewidth: width,
            worldUnits: false,
            transparent: false
        });
        material.resolution.set(this.canvas.clientWidth, this.canvas.clientHeight);
        this.constructionLineMaterials.add(material);

        const line = new Line2(geometry, material);
        line.computeLineDistances();
        return line;
    }

    createTriangle(a, b, c, color, opacity) {
        const geometry = new THREE.BufferGeometry();
        const vertices = new Float32Array([
            a.x, a.y, a.z,
            b.x, b.y, b.z,
            c.x, c.y, c.z
        ]);
        geometry.setAttribute('position', new THREE.BufferAttribute(vertices, 3));
        geometry.setIndex([0, 1, 2]);
        geometry.computeVertexNormals();

        const material = new THREE.MeshBasicMaterial({
            color,
            transparent: true,
            opacity,
            side: THREE.DoubleSide,
            depthTest: false,
            depthWrite: false,
            polygonOffset: true,
            polygonOffsetFactor: -2,
            polygonOffsetUnits: -2
        });
        const mesh = new THREE.Mesh(geometry, material);
        mesh.renderOrder = 20;

        const outline = this.createSegment(a, b, color);
        const outline2 = this.createSegment(b, c, color);
        const outline3 = this.createSegment(c, a, color);
        outline.renderOrder = 21;
        outline2.renderOrder = 21;
        outline3.renderOrder = 21;

        const rightAngleMarkers = [
            this.createRightAngleMarker(a, b, c),
            this.createRightAngleMarker(b, c, a),
            this.createRightAngleMarker(c, a, b)
        ].filter(Boolean);

        const group = new THREE.Group();
        group.add(mesh, outline, outline2, outline3, ...rightAngleMarkers);
        return group;
    }

    createRightAngleMarker(vertex, armPoint1, armPoint2) {
        const arm1 = armPoint1.clone().sub(vertex);
        const arm2 = armPoint2.clone().sub(vertex);
        const len1 = arm1.length();
        const len2 = arm2.length();
        if (len1 < 1e-6 || len2 < 1e-6) {
            return null;
        }

        const u = arm1.clone().normalize();
        const v = arm2.clone().normalize();
        const cosTheta = THREE.MathUtils.clamp(u.dot(v), -1, 1);
        const angle = Math.acos(cosTheta);
        const rightAngleToleranceRadians = THREE.MathUtils.degToRad(2);
        if (Math.abs(angle - Math.PI / 2) > rightAngleToleranceRadians) {
            return null;
        }

        const markerSize = THREE.MathUtils.clamp(Math.min(len1, len2) * 0.16, 0.15, 0.9);
        const p1 = vertex.clone().add(u.clone().multiplyScalar(markerSize));
        const p2 = p1.clone().add(v.clone().multiplyScalar(markerSize));
        const p3 = vertex.clone().add(v.clone().multiplyScalar(markerSize));
        const markerLine = this.createThickPolyline([p1, p2, p3], 0x000000, 4.5);
        markerLine.renderOrder = 22;
        return markerLine;
    }

    createQuad(points, color, opacity) {
        const geometry = new THREE.BufferGeometry();
        const vertices = new Float32Array(points.flatMap((point) => [point.x, point.y, point.z]));
        geometry.setAttribute('position', new THREE.BufferAttribute(vertices, 3));
        geometry.setIndex([0, 1, 2, 0, 2, 3]);
        geometry.computeVertexNormals();
        const mesh = new THREE.Mesh(
            geometry,
            new THREE.MeshBasicMaterial({
                color,
                transparent: true,
                opacity,
                side: THREE.DoubleSide,
                depthTest: false,
                depthWrite: false,
                polygonOffset: true,
                polygonOffsetFactor: -2,
                polygonOffsetUnits: -2
            })
        );
        mesh.renderOrder = 20;

        const outline = new THREE.Group();
        outline.add(mesh);
        for (let index = 0; index < points.length; index += 1) {
            const nextIndex = (index + 1) % points.length;
            const edge = this.createSegment(points[index], points[nextIndex], color);
            edge.renderOrder = 21;
            outline.add(edge);
        }
        return outline;
    }

    createAngleMarker(a, vertex, c, angleText, arcColor = 0x00d1b2) {
        const radius = Math.min(a.distanceTo(vertex), c.distanceTo(vertex)) * 0.28;
        const dir1 = a.clone().sub(vertex).normalize();
        const dir2 = c.clone().sub(vertex).normalize();
        const normal = new THREE.Vector3().crossVectors(dir1, dir2).normalize();
        const tangent = new THREE.Vector3().crossVectors(normal, dir1).normalize();
        const rawAngle = Math.acos(THREE.MathUtils.clamp(dir1.dot(dir2), -1, 1));
        const sampleCount = 28;
        const points = [];

        for (let step = 0; step <= sampleCount; step += 1) {
            const theta = (rawAngle * step) / sampleCount;
            const arcPoint = vertex.clone()
                .add(dir1.clone().multiplyScalar(Math.cos(theta) * radius))
                .add(tangent.clone().multiplyScalar(Math.sin(theta) * radius));
            points.push(arcPoint);
        }

        const line = this.createThickPolyline(points, arcColor, 5);

        const labelRadius = radius * 0.72;
        const labelPoint = vertex.clone()
            .add(dir1.clone().multiplyScalar(Math.cos(rawAngle / 2) * labelRadius))
            .add(tangent.clone().multiplyScalar(Math.sin(rawAngle / 2) * labelRadius));
        const label = this.createTextSprite(`∠${angleText}`, {
            fontSize: 42,
            textColor: '#000000',
            background: '#9de7ff',
            borderColor: '#000000'
        });
        label.position.copy(labelPoint);

        const group = new THREE.Group();
        group.add(line, label);
        return group;
    }

    createTextSprite(text, options = {}) {
        const fontSize = options.fontSize || 48;
        const padding = 22;
        const dpr = Math.min(window.devicePixelRatio || 1, 2);
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        context.font = `700 ${fontSize}px Segoe UI`;
        const metrics = context.measureText(text);
        const logicalWidth = Math.ceil(metrics.width + padding * 2);
        const logicalHeight = Math.ceil(fontSize + padding * 2);
        canvas.width = Math.ceil(logicalWidth * dpr);
        canvas.height = Math.ceil(logicalHeight * dpr);

        const drawContext = canvas.getContext('2d');
        drawContext.setTransform(dpr, 0, 0, dpr, 0, 0);
        drawContext.clearRect(0, 0, logicalWidth, logicalHeight);
        drawContext.font = `700 ${fontSize}px Segoe UI`;
        drawContext.textBaseline = 'middle';
        drawContext.textAlign = 'center';
        drawContext.lineJoin = 'round';
        drawContext.lineCap = 'round';

        const background = options.background || '#ffd84d';
        const borderColor = options.borderColor || '#000000';
        const borderWidth = options.borderWidth || 4;
        const radius = options.cornerRadius || 16;
        drawContext.fillStyle = background;
        this.drawRoundedRect(drawContext, borderWidth / 2, borderWidth / 2, logicalWidth - borderWidth, logicalHeight - borderWidth, radius);
        drawContext.fill();

        drawContext.lineWidth = borderWidth;
        drawContext.strokeStyle = borderColor;
        drawContext.stroke();

        drawContext.fillStyle = options.textColor || '#ffffff';
        drawContext.fillText(text, logicalWidth / 2, logicalHeight / 2 + 2);

        const texture = new THREE.CanvasTexture(canvas);
        texture.needsUpdate = true;
        texture.generateMipmaps = false;
        texture.minFilter = THREE.LinearFilter;
        texture.magFilter = THREE.LinearFilter;
        const material = new THREE.SpriteMaterial({ map: texture, transparent: true });
        material.alphaTest = 0.08;
        material.depthTest = true;
        material.depthWrite = false;
        const sprite = new THREE.Sprite(material);
        const scaleFactor = 0.0055;
        sprite.scale.set(logicalWidth * scaleFactor, logicalHeight * scaleFactor, 1);
        sprite.renderOrder = 40;
        return sprite;
    }

    drawRoundedRect(context, x, y, width, height, radius) {
        context.beginPath();
        context.moveTo(x + radius, y);
        context.lineTo(x + width - radius, y);
        context.quadraticCurveTo(x + width, y, x + width, y + radius);
        context.lineTo(x + width, y + height - radius);
        context.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        context.lineTo(x + radius, y + height);
        context.quadraticCurveTo(x, y + height, x, y + height - radius);
        context.lineTo(x, y + radius);
        context.quadraticCurveTo(x, y, x + radius, y);
        context.closePath();
    }

    addSceneObject({ type, name, subtitle, object3D, definition = null }) {
        object3D.userData.sceneObjectId = this.nextObjectId;
        this.scene.add(object3D);
        this.sceneObjects.unshift({
            id: this.nextObjectId,
            type,
            name,
            subtitle,
            object3D,
            definition,
            visible: true
        });
        this.nextObjectId += 1;
        this.renderObjectsList();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    getVectorsByPointIds(pointIds) {
        const points = pointIds.map((pointId) => this.getPointById(pointId));
        if (points.some((point) => !point)) {
            return null;
        }
        return points.map((point) => point.position.clone());
    }

    rebuildConstructions() {
        if (this.sceneObjects.length === 0) {
            return;
        }

        const survivingObjects = [];
        this.sceneObjects.forEach((entry) => {
            if (!entry.definition) {
                survivingObjects.push(entry);
                return;
            }

            const rebuiltObject = this.createObjectFromDefinition(entry.definition);
            if (!rebuiltObject) {
                if (entry.object3D) {
                    this.scene.remove(entry.object3D);
                    this.disposeObject3D(entry.object3D);
                }
                return;
            }

            if (entry.object3D) {
                this.scene.remove(entry.object3D);
                this.disposeObject3D(entry.object3D);
            }

            rebuiltObject.userData.sceneObjectId = entry.id;
            rebuiltObject.visible = entry.visible;
            entry.object3D = rebuiltObject;
            this.scene.add(rebuiltObject);
            survivingObjects.push(entry);
        });

        this.sceneObjects = survivingObjects;
        this.renderObjectsList();
    }

    createObjectFromDefinition(definition) {
        if (!definition || !definition.kind) {
            return null;
        }

        if (definition.kind === 'point-label') {
            const point = this.getPointById(definition.pointId);
            if (!point) return null;
            const sprite = this.createTextSprite(definition.text, {
                fontSize: 46,
                textColor: '#000000',
                background: '#ffd84d',
                borderColor: '#000000'
            });
            sprite.position.copy(point.position.clone().add(new THREE.Vector3(0.34, 0.46, 0.18)));
            return sprite;
        }

        if (definition.kind === 'length-label') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 2) return null;
            const sprite = this.createTextSprite(definition.text, {
                fontSize: 46,
                textColor: '#000000',
                background: '#b9f18a',
                borderColor: '#000000'
            });
            const midpoint = vectors[0].clone().lerp(vectors[1], 0.5).add(new THREE.Vector3(0.2, 0.2, 0.2));
            sprite.position.copy(midpoint);
            return sprite;
        }

        if (definition.kind === 'segment') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 2) return null;
            return this.createSegment(vectors[0], vectors[1], definition.color || 0xff595e);
        }

        if (definition.kind === 'triangle') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 3) return null;
            return this.createTriangle(vectors[0], vectors[1], vectors[2], definition.color || 0xff595e, definition.opacity || 0.28);
        }

        if (definition.kind === 'angle') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 3) return null;
            return this.createAngleMarker(vectors[0], vectors[1], vectors[2], definition.pointIds.join(''), definition.color || 0xff595e);
        }

        if (definition.kind === 'plane') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 4) return null;
            return this.createQuad(vectors, definition.color || 0xff595e, definition.opacity || 0.2);
        }

        return null;
    }

    renderObjectsList() {
        const SECTION_TYPES = { triangles: 'triangle', segments: 'segment', angles: 'angle', planes: 'plane', labels: 'label' };
        for (const [key, type] of Object.entries(SECTION_TYPES)) {
            const sec = this.objectSections[key];
            const items = this.sceneObjects.filter((item) => item.type === type);
            sec.list.innerHTML = '';
            items.forEach((item) => sec.list.appendChild(this.renderObjectItem(item)));
            sec.count.textContent = items.length > 0 ? `(${items.length})` : '';
        }
    }



    renderObjectItem(item) {
        const row = document.createElement('div');
        row.className = 'object-item';
        row.innerHTML = `
            <div class="object-name">
                <strong>${item.name}</strong>
                <span>${item.subtitle}</span>
            </div>
            <button type="button" class="object-toggle ${item.visible ? '' : 'is-hidden'}" data-toggle-object-id="${item.id}" aria-label="Toggle object visibility">${item.visible ? '◐' : '◌'}</button>
            <button type="button" class="object-delete" data-delete-object-id="${item.id}" aria-label="Delete object">×</button>
        `;
        return row;
    }

    toggleObjectVisibility(objectId) {
        const item = this.sceneObjects.find((entry) => entry.id === objectId);
        if (!item) return;

        item.visible = !item.visible;
        item.object3D.visible = item.visible;
        this.renderObjectsList();
    }

    deleteObject(objectId) {
        const index = this.sceneObjects.findIndex((entry) => entry.id === objectId);
        if (index === -1) return;

        const [item] = this.sceneObjects.splice(index, 1);
        this.scene.remove(item.object3D);
        this.disposeObject3D(item.object3D);
        this.renderObjectsList();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    clearObjects() {
        this.sceneObjects.forEach((item) => {
            this.scene.remove(item.object3D);
            this.disposeObject3D(item.object3D);
        });
        this.sceneObjects = [];
        this.renderObjectsList();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    resetSceneObjects() {
        this.clearObjects();
        this.clearSelection();
    }

    fitCameraToPrimitive(radius) {
        const distance = Math.max(radius * 2.1, 8);
        this.camera.position.set(distance, distance * 0.72, distance * 0.94);
        this.controls.target.set(0, 0, 0);
        this.controls.update();
    }

    resetView() {
        this.fitCameraToPrimitive(6);
    }

    clearComposite() {
        if (this.compositeGroup) {
            this.scene.remove(this.compositeGroup);
            this.disposeObject3D(this.compositeGroup);
            this.compositeGroup = null;
            this.primitiveGroup = null;
        }
        this.slotGroupMap = new Map();
        this.primitiveMeshes = [];
        this.slotLinkages = [];
        this.pointMarkers.clear();
        this.pointDefinitions = [];
        this.derivedPoints = [];
    }

    clearPrimitive() {
        this.clearComposite();
    }

    disposeSprites(sprites) {
        sprites.forEach((sprite) => this.disposeObject3D(sprite));
    }

    disposeObject3D(object3D) {
        object3D.traverse?.((child) => {
            if (child.geometry) {
                child.geometry.dispose();
            }
            if (child.material) {
                if (Array.isArray(child.material)) {
                    child.material.forEach((material) => {
                        material.map?.dispose?.();
                        if (this.constructionLineMaterials.has(material)) {
                            this.constructionLineMaterials.delete(material);
                        }
                        material.dispose();
                    });
                } else {
                    child.material.map?.dispose?.();
                    if (this.constructionLineMaterials.has(child.material)) {
                        this.constructionLineMaterials.delete(child.material);
                    }
                    child.material.dispose();
                }
            }
        });
    }

    animate() {
        if (!this.renderer) return;
        this.animationFrameId = window.requestAnimationFrame(() => this.animate());
        this.controls.update();
        this.renderer.render(this.scene, this.camera);
    }

    cleanup() {
        window.removeEventListener('resize', this.handleWindowResize);
        this.canvas.removeEventListener('pointerdown', this.handleCanvasPointerDown);
        this.clearObjects();
        this.clearPrimitive();
        window.cancelAnimationFrame(this.animationFrameId);
        this.controls.dispose();
        this.renderer.dispose();
    }
}