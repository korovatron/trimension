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

function normalizeTriangularPrismMode(mode) {
    return mode === 'equilateral' ? 'isosceles' : (mode || 'isosceles');
}

function getTriangularPrismProfilePoints(params, zPos) {
    const { legA, legB } = params;
    const mode = normalizeTriangularPrismMode(params.triangleMode);

    if (mode === 'isosceles') {
        return [
            new THREE.Vector3(-legA / 2, -legB / 2, zPos),
            new THREE.Vector3(legA / 2, -legB / 2, zPos),
            new THREE.Vector3(0, legB / 2, zPos)
        ];
    }

    if (mode === 'right-above-B') {
        return [
            new THREE.Vector3(0, -legB / 2, zPos),
            new THREE.Vector3(legA, -legB / 2, zPos),
            new THREE.Vector3(legA, legB / 2, zPos)
        ];
    }

    return [
        new THREE.Vector3(0, -legB / 2, zPos),
        new THREE.Vector3(legA, -legB / 2, zPos),
        new THREE.Vector3(0, legB / 2, zPos)
    ];
}

function getTriangleCentroid(a, b, c) {
    return new THREE.Vector3(
        (a.x + b.x + c.x) / 3,
        (a.y + b.y + c.y) / 3,
        (a.z + b.z + c.z) / 3
    );
}

function normalizeTetrahedronBaseMode(mode) {
    return mode || 'isosceles';
}

function getTetrahedronBasePoints(params, yPos) {
    const { base, triangleHeight } = params;
    const mode = normalizeTetrahedronBaseMode(params.baseTriangleMode);

    if (mode === 'right-angled') {
        return [
            new THREE.Vector3(-base / 3, yPos, -triangleHeight / 3),
            new THREE.Vector3((2 * base) / 3, yPos, -triangleHeight / 3),
            new THREE.Vector3(-base / 3, yPos, (2 * triangleHeight) / 3)
        ];
    }

    return [
        new THREE.Vector3(-base / 2, yPos, -triangleHeight / 3),
        new THREE.Vector3(base / 2, yPos, -triangleHeight / 3),
        new THREE.Vector3(0, yPos, (2 * triangleHeight) / 3)
    ];
}

const ATTACHMENT_FACES = {
    cuboid: [
        { id: 'top',    type: 'rectangle', normal: new THREE.Vector3(0, 1, 0),   uAxis: new THREE.Vector3(1, 0, 0), center: (p) => new THREE.Vector3(0, p.height / 2, 0),   dims: ['width', 'depth'],  label: 'Top' },
        { id: 'bottom', type: 'rectangle', normal: new THREE.Vector3(0, -1, 0),  uAxis: new THREE.Vector3(1, 0, 0), center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['width', 'depth'],  label: 'Bottom' },
        { id: 'front',  type: 'rectangle', normal: new THREE.Vector3(0, 0, 1),   uAxis: new THREE.Vector3(1, 0, 0), center: (p) => new THREE.Vector3(0, 0, p.depth / 2),     dims: ['width', 'height'], label: 'Front' },
        { id: 'back',   type: 'rectangle', normal: new THREE.Vector3(0, 0, -1),  uAxis: new THREE.Vector3(1, 0, 0), center: (p) => new THREE.Vector3(0, 0, -p.depth / 2),    dims: ['width', 'height'], label: 'Back' },
        { id: 'right',  type: 'rectangle', normal: new THREE.Vector3(1, 0, 0),   uAxis: new THREE.Vector3(0, 0, 1), center: (p) => new THREE.Vector3(p.width / 2, 0, 0),     dims: ['depth', 'height'], label: 'Right' },
        { id: 'left',   type: 'rectangle', normal: new THREE.Vector3(-1, 0, 0),  uAxis: new THREE.Vector3(0, 0, 1), center: (p) => new THREE.Vector3(-p.width / 2, 0, 0),    dims: ['depth', 'height'], label: 'Left' },
    ],
    'rectangular-pyramid': [
        { id: 'base',   type: 'rectangle', normal: new THREE.Vector3(0, -1, 0),  uAxis: new THREE.Vector3(1, 0, 0), center: (p) => new THREE.Vector3(0, -p.height / 2, 0),  dims: ['length', 'width'], label: 'Base' },
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
    'right-triangle-prism': [
        {
            id: 'base-rectangle',
            type: 'rectangle',
            normal: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(0, -1, 0);
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(0, -1, 0);
                } else {
                    return new THREE.Vector3(0, -1, 0);
                }
            },
            uAxis: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(1, 0, 0);
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(1, 0, 0);
                } else {
                    return new THREE.Vector3(1, 0, 0);
                }
            },
            center: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(0, -p.legB / 2, 0);
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(p.legA / 2, -p.legB / 2, 0);
                } else {
                    return new THREE.Vector3(p.legA / 2, -p.legB / 2, 0);
                }
            },
            dims: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return ['legA', 'length'];
                } else {
                    return ['legA', 'length'];
                }
            },
            label: 'Base Rectangle'
        },
        {
            id: 'side-rectangle',
            type: 'rectangle',
            normal: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(-p.legB, p.legA / 2, 0).normalize();
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(1, 0, 0);
                } else {
                    return new THREE.Vector3(-1, 0, 0);
                }
            },
            uAxis: new THREE.Vector3(0, 0, 1),
            center: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(-p.legA / 4, 0, 0);
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(p.legA, 0, 0);
                } else {
                    return new THREE.Vector3(0, 0, 0);
                }
            },
            dims: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return ['length', 'isoscelesSide'];
                } else {
                    return ['length', 'legB'];
                }
            },
            label: 'Side Rectangle'
        },
        {
            id: 'hypotenuse-rectangle',
            type: 'rectangle',
            normal: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(p.legB, p.legA / 2, 0).normalize();
                } else if (mode === 'right-above-B') {
                    return new THREE.Vector3(-p.legB, p.legA, 0).normalize();
                } else {
                    return new THREE.Vector3(p.legB, p.legA, 0).normalize();
                }
            },
            uAxis: new THREE.Vector3(0, 0, 1),
            center: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return new THREE.Vector3(p.legA / 4, 0, 0);
                } else {
                    return new THREE.Vector3(p.legA / 2, 0, 0);
                }
            },
            dims: (p) => {
                const mode = p.triangleMode === 'equilateral' ? 'isosceles' : (p.triangleMode || 'isosceles');
                if (mode === 'isosceles') {
                    return ['length', 'isoscelesSide'];
                } else {
                    return ['length', 'hypotenuse'];
                }
            },
            label: 'Hypotenuse Rectangle'
        },
        {
            id: 'front-triangle',
            type: 'triangle',
            normal: new THREE.Vector3(0, 0, 1),
            uAxis: new THREE.Vector3(1, 0, 0),
            center: (p) => {
                const [a, b, c] = getTriangularPrismProfilePoints(p, p.length / 2);
                return getTriangleCentroid(a, b, c);
            },
            dims: ['legA', 'legB'],
            label: 'Front Triangle'
        },
        {
            id: 'back-triangle',
            type: 'triangle',
            normal: new THREE.Vector3(0, 0, -1),
            uAxis: new THREE.Vector3(1, 0, 0),
            center: (p) => {
                const [a, b, c] = getTriangularPrismProfilePoints(p, -p.length / 2);
                return getTriangleCentroid(a, b, c);
            },
            dims: ['legA', 'legB'],
            label: 'Back Triangle'
        },
    ],
    tetrahedron: [
        {
            id: 'base-triangle',
            type: 'triangle',
            normal: new THREE.Vector3(0, -1, 0),
            uAxis: new THREE.Vector3(1, 0, 0),
            center: (p) => new THREE.Vector3(0, -p.height / 2, 0),
            dims: ['base', 'triangleHeight'],
            label: 'Base Triangle'
        },
    ],
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
        this.derivedLabelOverrides = new Map();
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
        this.primitiveSectionCollapsed = false;
        this.pointsSectionCollapsed = true;

        this.defaultParams = {
            cuboid: { width: 7, depth: 4, height: 5 },
            'right-triangle-prism': { legA: 5, legB: 4, length: 7, triangleMode: 'isosceles' },
            tetrahedron: { base: 6, triangleHeight: 4.5, height: 6, baseTriangleMode: 'isosceles', apexPosition: 'A' },
            'trapezium-prism': { baseWidth: 6, leftHeight: 4, rightHeight: 2.5, length: 7 },
            sphere: { radius: 3 },
            hemisphere: { radius: 3 },
            cylinder: { radius: 2.5, height: 6 },
            cone: { radius: 2.5, height: 6 },
            'rectangular-pyramid': { length: 6.5, width: 4.5, height: 6, apexPosition: 'center' }
        };

        // compositeSlots: array of { id, primitive, orientation, params, hostSlotId, hostFaceId }
        this.compositeSlots = [
            { id: 0, primitive: 'cuboid', orientation: 'standard', params: { width: 7, depth: 4, height: 5 }, hostSlotId: null, hostFaceId: null }
        ];
        this.nextSlotId = 1;
        this.compositeGroup = null;
        this.slotGroupMap = new Map();   // slotId -> Three.js Group
        this.primitiveMeshes = [];       // one mesh per slot
        this.slotLinkages = [];          // { fromSlotId, fromParam, toSlotId, toParam }

        this.orientations = {
            cuboid: [
                { value: 'standard', label: 'Standard' }
            ],
            'right-triangle-prism': [
                { value: 'standard', label: 'Standard' }
            ],
            tetrahedron: [
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

        this.rectangularPyramidApexPositions = [
            { value: 'center', label: 'Centre' },
            { value: 'A', label: 'Above A' },
            { value: 'B', label: 'Above B' },
            { value: 'C', label: 'Above C' },
            { value: 'D', label: 'Above D' }
        ];

        this.triangularPrismModes = [
            { value: 'isosceles', label: 'Isosceles' },
            { value: 'right-above-A', label: 'Right Angle at A' },
            { value: 'right-above-B', label: 'Right Angle at B' }
        ];

        this.tetrahedronTriangleModes = [
            { value: 'isosceles', label: 'Isosceles' },
            { value: 'right-angled', label: 'Right-Angled' }
        ];

        this.tetrahedronApexPositions = [
            { value: 'A', label: 'Above A' },
            { value: 'B', label: 'Above B' },
            { value: 'C', label: 'Above C' }
        ];

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
                label: 'Triangle Prism',
                params: [
                    { key: 'legA', label: 'Triangle Leg A', min: 2, max: 10, step: 0.5 },
                    { key: 'legB', label: 'Triangle Leg B', min: 2, max: 10, step: 0.5 },
                    { key: 'length', label: 'Prism Length', min: 2, max: 12, step: 0.5 }
                ]
            },
            tetrahedron: {
                label: 'Tetrahedron',
                params: [
                    { key: 'base', label: 'Triangle Base', min: 2, max: 10, step: 0.5 },
                    { key: 'triangleHeight', label: 'Triangle Height', min: 2, max: 10, step: 0.5 },
                    { key: 'height', label: 'Apex Height', min: 2, max: 10, step: 0.5 }
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
                { key: 'change-point-label', label: 'Change Label' }
            ],
            2: [
                { key: 'segment', label: 'Add Segment' }
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

            const triangleModeBtn = event.target.closest('[data-cycle-triangle-mode-slot-id]');
            if (triangleModeBtn) {
                this.cycleTriangularPrismMode(Number(triangleModeBtn.dataset.cycleTriangleModeSlotId));
                return;
            }

            const tetrahedronModeBtn = event.target.closest('[data-cycle-tetrahedron-mode-slot-id]');
            if (tetrahedronModeBtn) {
                this.cycleTetrahedronTriangleMode(Number(tetrahedronModeBtn.dataset.cycleTetrahedronModeSlotId));
                return;
            }

            const tetrahedronApexBtn = event.target.closest('[data-cycle-tetrahedron-apex-slot-id]');
            if (tetrahedronApexBtn) {
                this.cycleTetrahedronApex(Number(tetrahedronApexBtn.dataset.cycleTetrahedronApexSlotId));
                return;
            }

            const apexCycleBtn = event.target.closest('[data-cycle-apex-slot-id]');
            if (apexCycleBtn) {
                this.cycleRectangularPyramidApex(Number(apexCycleBtn.dataset.cycleApexSlotId));
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

        this.pointsSectionHeader.addEventListener('click', togglePointsSection);

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
        const disallowedSelfAdds = new Set(['cuboid', 'cylinder', 'hemisphere']);

        return Object.keys(this.defaultParams).filter((primKey) => {
            if (existingPrimitives.has(primKey) && disallowedSelfAdds.has(primKey)) return false;
            const guestFaces = ATTACHMENT_FACES[primKey] || [];
            if (guestFaces.length === 0) return false;
            return guestFaces.some((gf) => existingFaceTypes.has(gf.type));
        });
    }

    getHostFaceKey(slotId, faceId) {
        return `${slotId}:${faceId}`;
    }

    getFaceDefById(slot, faceId) {
        if (!slot || !faceId) return null;
        return (ATTACHMENT_FACES[slot.primitive] || []).find((f) => f.id === faceId) || null;
    }

    resolveFaceNormal(faceDef, params) {
        const raw = typeof faceDef.normal === 'function' ? faceDef.normal(params) : faceDef.normal;
        const normal = raw ? raw.clone() : new THREE.Vector3(0, 1, 0);
        if (normal.lengthSq() < 1e-8) return new THREE.Vector3(0, 1, 0);
        return normal.normalize();
    }

    resolveFaceUAxis(faceDef, params) {
        const raw = typeof faceDef.uAxis === 'function' ? faceDef.uAxis(params) : faceDef.uAxis;
        const uAxis = raw ? raw.clone() : new THREE.Vector3(1, 0, 0);
        if (uAxis.lengthSq() < 1e-8) return new THREE.Vector3(1, 0, 0);
        return uAxis.normalize();
    }

    resolveFaceDims(faceDef, params) {
        const raw = typeof faceDef.dims === 'function' ? faceDef.dims(params) : faceDef.dims;
        return Array.isArray(raw) ? raw : [];
    }

    getFaceDimValue(slot, faceDef, dimKey) {
        if (!slot || !faceDef || !dimKey) return undefined;
        if (Object.prototype.hasOwnProperty.call(slot.params, dimKey)) {
            return slot.params[dimKey];
        }

        if (slot.primitive === 'right-triangle-prism' && dimKey === 'isoscelesSide') {
            const { legA, legB } = slot.params;
            return Math.hypot(legA / 2, legB);
        }

        if (slot.primitive === 'right-triangle-prism' && dimKey === 'hypotenuse') {
            const { legA, legB } = slot.params;
            return Math.hypot(legA, legB);
        }

        return undefined;
    }

    getOccupiedHostFaceKeys(excludeSlotId = null) {
        const occupied = new Set();
        this.compositeSlots.forEach((slot) => {
            if (slot.id === excludeSlotId) return;
            if (slot.hostSlotId == null || !slot.hostFaceId) return;
            occupied.add(this.getHostFaceKey(slot.hostSlotId, slot.hostFaceId));

            // Also block the guest face at this join so internal faces cannot be selected.
            const hostSlot = this.compositeSlots.find((s) => s.id === slot.hostSlotId);
            const hostFaceDef = this.getFaceDefById(hostSlot, slot.hostFaceId);
            if (!hostFaceDef) return;

            const hostFaceNormal = this.resolveFaceNormal(hostFaceDef, hostSlot.params);
            const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal);
            if (guestFaceDef) {
                occupied.add(this.getHostFaceKey(slot.id, guestFaceDef.id));
            }
        });
        return occupied;
    }

    getValidHostFaceEntries(guestSlot, previousSlots, options = {}) {
        const excludeOccupied = options.excludeOccupied === true;
        const excludeSlotId = options.excludeSlotId ?? null;
        const guestFaceTypes = new Set((ATTACHMENT_FACES[guestSlot.primitive] || []).map((f) => f.type));
        if (guestFaceTypes.size === 0) return [];

        const occupiedKeys = excludeOccupied ? this.getOccupiedHostFaceKeys(excludeSlotId) : null;

        const result = [];
        previousSlots.forEach((hostSlot) => {
            (ATTACHMENT_FACES[hostSlot.primitive] || []).forEach((faceDef) => {
                if (guestFaceTypes.has(faceDef.type)) {
                    const key = this.getHostFaceKey(hostSlot.id, faceDef.id);
                    if (occupiedKeys && occupiedKeys.has(key)) {
                        return;
                    }
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

    ensureSlotHostBinding(slot, previousSlots) {
        const allEntries = this.getValidHostFaceEntries(slot, previousSlots);
        if (allEntries.length === 0) {
            slot.hostSlotId = null;
            slot.hostFaceId = null;
            return null;
        }

        const occupiedByOthers = this.getOccupiedHostFaceKeys(slot.id);
        const findCurrent = () => allEntries.find((entry) => entry.slotId === slot.hostSlotId && entry.faceId === slot.hostFaceId);

        const currentEntry = findCurrent();
        if (currentEntry && !occupiedByOthers.has(this.getHostFaceKey(currentEntry.slotId, currentEntry.faceId))) {
            return currentEntry;
        }

        const availableEntries = allEntries.filter((entry) => !occupiedByOthers.has(this.getHostFaceKey(entry.slotId, entry.faceId)));
        const picked = availableEntries[0] || currentEntry || allEntries[0];
        slot.hostSlotId = picked.slotId;
        slot.hostFaceId = picked.faceId;
        return picked;
    }

    getGuestAttachFaceDef(guestSlot, hostFaceNormal) {
        const faces = ATTACHMENT_FACES[guestSlot.primitive] || [];
        if (faces.length === 0) return null;

        const targetNormal = hostFaceNormal.clone().negate();
        const ABS_DOT_EPSILON = 1e-6;
        const DOT_EPSILON = 1e-6;
        let best = null;
        let bestIndex = Infinity;
        let bestAbsDot = -Infinity;
        let bestDot = -Infinity;

        faces.forEach((face, index) => {
            const faceNormal = this.resolveFaceNormal(face, guestSlot.params);
            const dot = faceNormal.dot(targetNormal);
            const absDot = Math.abs(dot);

            if (best === null) {
                best = face;
                bestIndex = index;
                bestAbsDot = absDot;
                bestDot = dot;
                return;
            }

            const absDelta = absDot - bestAbsDot;
            if (absDelta > ABS_DOT_EPSILON) {
                best = face;
                bestIndex = index;
                bestAbsDot = absDot;
                bestDot = dot;
                return;
            }

            if (Math.abs(absDelta) <= ABS_DOT_EPSILON) {
                const dotDelta = dot - bestDot;
                if (dotDelta > DOT_EPSILON || (Math.abs(dotDelta) <= DOT_EPSILON && index < bestIndex)) {
                    best = face;
                    bestIndex = index;
                    bestAbsDot = absDot;
                    bestDot = dot;
                }
            }
        });
        return best;
    }

    snapSlotDimensions(guestSlot) {
        const guestSlotIndex = this.compositeSlots.findIndex((s) => s.id === guestSlot.id);
        if (guestSlotIndex <= 0) return;
        const prevSlots = this.compositeSlots.slice(0, guestSlotIndex);
        const entry = this.ensureSlotHostBinding(guestSlot, prevSlots);
        if (!entry) return;

        const { slot: hostSlot, faceDef: hostFaceDef } = entry;
        const hostFaceNormal = this.resolveFaceNormal(hostFaceDef, hostSlot.params);
        const guestFaceDef = this.getGuestAttachFaceDef(guestSlot, hostFaceNormal);
        if (!guestFaceDef) return;

        const hostDims = this.resolveFaceDims(hostFaceDef, hostSlot.params);
        const guestDims = this.resolveFaceDims(guestFaceDef, guestSlot.params);
        hostDims.forEach((dim, i) => {
            const guestDim = guestDims[i];
            const hostValue = this.getFaceDimValue(hostSlot, hostFaceDef, dim);
            if (guestSlot.primitive === 'right-triangle-prism' && guestDim === 'isoscelesSide') {
                const halfBase = guestSlot.params.legA / 2;
                if (typeof hostValue === 'number' && Number.isFinite(hostValue) && hostValue >= halfBase) {
                    guestSlot.params.legB = Math.sqrt(Math.max(0, hostValue * hostValue - halfBase * halfBase));
                }
                return;
            }
            if (!guestDim || !Object.prototype.hasOwnProperty.call(guestSlot.params, guestDim)) return;
            if (typeof hostValue === 'number' && Number.isFinite(hostValue)) {
                guestSlot.params[guestDim] = hostValue;
            }
        });
    }

    addSlot(primitiveKey) {
        if (!this.primitiveMeta[primitiveKey]) return;

        if (this.compositeSlots.length === 0) {
            // First slot - reset everything
            const slot = {
                id: this.nextSlotId++,
                primitive: primitiveKey,
                orientation: this.orientations[primitiveKey][0].value,
                params: { ...this.defaultParams[primitiveKey] },
                hostSlotId: null,
                hostFaceId: null,
            };
            this.compositeSlots.push(slot);
        } else {
            if (this.compositeSlots.length >= 3) return;
            const slot = {
                id: this.nextSlotId++,
                primitive: primitiveKey,
                orientation: this.orientations[primitiveKey][0].value,
                params: { ...this.defaultParams[primitiveKey] },
                hostSlotId: null,
                hostFaceId: null,
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
        // Remove only this slot; remaining slots will be re-evaluated against available hosts.
        this.compositeSlots.splice(idx, 1);
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    cycleSlotFace(slotId) {
        const slotIdx = this.compositeSlots.findIndex((s) => s.id === slotId);
        if (slotIdx <= 0) return;
        const slot = this.compositeSlots[slotIdx];
        const prevSlots = this.compositeSlots.slice(0, slotIdx);
        const entries = this.getValidHostFaceEntries(slot, prevSlots, { excludeOccupied: true, excludeSlotId: slot.id });
        if (entries.length <= 1) return;

        const currentIndex = entries.findIndex((entry) => entry.slotId === slot.hostSlotId && entry.faceId === slot.hostFaceId);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % entries.length : 0;
        slot.hostSlotId = entries[nextIndex].slotId;
        slot.hostFaceId = entries[nextIndex].faceId;
        this.snapSlotDimensions(slot);
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getTriangularPrismModeLabel(mode) {
        const normalizedMode = mode === 'equilateral' ? 'isosceles' : mode;
        return this.triangularPrismModes.find((opt) => opt.value === normalizedMode)?.label || 'Isosceles';
    }

    cycleTriangularPrismMode(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'right-triangle-prism') return;

        const options = this.triangularPrismModes;
        const current = slot.params.triangleMode === 'equilateral' ? 'isosceles' : (slot.params.triangleMode || 'isosceles');
        const currentIndex = options.findIndex((opt) => opt.value === current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % options.length : 0;
        slot.params.triangleMode = options[nextIndex].value;

        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getTetrahedronTriangleModeLabel(mode) {
        return this.tetrahedronTriangleModes.find((opt) => opt.value === normalizeTetrahedronBaseMode(mode))?.label || 'Isosceles';
    }

    cycleTetrahedronTriangleMode(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'tetrahedron') return;

        const options = this.tetrahedronTriangleModes;
        const current = normalizeTetrahedronBaseMode(slot.params.baseTriangleMode);
        const currentIndex = options.findIndex((opt) => opt.value === current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % options.length : 0;
        slot.params.baseTriangleMode = options[nextIndex].value;

        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getTetrahedronApexLabel(apexPosition) {
        return this.tetrahedronApexPositions.find((opt) => opt.value === apexPosition)?.label || 'Above A';
    }

    cycleTetrahedronApex(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'tetrahedron') return;

        const options = this.tetrahedronApexPositions;
        const current = slot.params.apexPosition || 'A';
        const currentIndex = options.findIndex((opt) => opt.value === current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % options.length : 0;
        slot.params.apexPosition = options[nextIndex].value;

        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getRectangularPyramidApexLabel(apexPosition) {
        return this.rectangularPyramidApexPositions.find((opt) => opt.value === apexPosition)?.label || 'Centre';
    }

    cycleRectangularPyramidApex(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'rectangular-pyramid') return;

        const options = this.rectangularPyramidApexPositions;
        const current = slot.params.apexPosition || 'center';
        const currentIndex = options.findIndex((opt) => opt.value === current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % options.length : 0;
        slot.params.apexPosition = options[nextIndex].value;

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

            const removeBtn = document.createElement('button');
            removeBtn.type = 'button';
            removeBtn.className = 'card-remove-btn';
            removeBtn.dataset.removeSlotId = String(slot.id);
            removeBtn.setAttribute('aria-label', `Remove ${this.primitiveMeta[slot.primitive].label}`);
            removeBtn.textContent = 'X';
            header.appendChild(removeBtn);
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
                this.updateSliderFill(input);

                input.addEventListener('input', () => {
                    const newVal = Number(input.value);
                    slot.params[config.key] = newVal;
                    this.updateSliderFill(input);

                    // Propagate to linked params
                    this.slotLinkages.forEach((link) => {
                        if (link.fromSlotId !== slot.id || link.fromParam !== config.key) return;
                        const linkedSlot = this.compositeSlots.find((s) => s.id === link.toSlotId);
                        if (!linkedSlot) return;
                        linkedSlot.params[link.toParam] = newVal;
                        const linkedRow = this.primitiveCardsListEl.querySelector(`[data-slot-id="${link.toSlotId}"] [data-param-key="${link.toParam}"]`);
                        if (linkedRow) {
                            const li = linkedRow.querySelector('.slider-input');
                            if (li) {
                                li.value = String(newVal);
                                this.updateSliderFill(li);
                            }
                        }
                    });

                    this.buildComposite({ fitCamera: false });
                });

                inputWrap.append(input);
                row.append(labelEl, inputWrap);
                sliderStack.appendChild(row);
            });

            card.appendChild(sliderStack);

            if (slot.primitive === 'right-triangle-prism') {
                const triangleModeCycleBtn = document.createElement('button');
                triangleModeCycleBtn.type = 'button';
                triangleModeCycleBtn.className = 'card-cycle-btn';
                triangleModeCycleBtn.dataset.cycleTriangleModeSlotId = String(slot.id);
                triangleModeCycleBtn.textContent = `Cycle: ${this.getTriangularPrismModeLabel(slot.params.triangleMode)}`;
                card.appendChild(triangleModeCycleBtn);
            }

            if (slot.primitive === 'tetrahedron') {
                const tetrahedronModeCycleBtn = document.createElement('button');
                tetrahedronModeCycleBtn.type = 'button';
                tetrahedronModeCycleBtn.className = 'card-cycle-btn';
                tetrahedronModeCycleBtn.dataset.cycleTetrahedronModeSlotId = String(slot.id);
                tetrahedronModeCycleBtn.textContent = `Cycle Base: ${this.getTetrahedronTriangleModeLabel(slot.params.baseTriangleMode)}`;
                card.appendChild(tetrahedronModeCycleBtn);

                const tetrahedronApexCycleBtn = document.createElement('button');
                tetrahedronApexCycleBtn.type = 'button';
                tetrahedronApexCycleBtn.className = 'card-cycle-btn';
                tetrahedronApexCycleBtn.dataset.cycleTetrahedronApexSlotId = String(slot.id);
                tetrahedronApexCycleBtn.textContent = `Cycle Apex: ${this.getTetrahedronApexLabel(slot.params.apexPosition)}`;
                card.appendChild(tetrahedronApexCycleBtn);
            }

            if (slot.primitive === 'rectangular-pyramid') {
                const apexCycleBtn = document.createElement('button');
                apexCycleBtn.type = 'button';
                apexCycleBtn.className = 'card-cycle-btn';
                apexCycleBtn.dataset.cycleApexSlotId = String(slot.id);
                apexCycleBtn.textContent = `Cycle Apex: ${this.getRectangularPyramidApexLabel(slot.params.apexPosition)}`;
                card.appendChild(apexCycleBtn);
            }

            // Cycle face button (slot 1+)
            if (idx > 0) {
                const prevSlots = this.compositeSlots.slice(0, idx);
                const entries = this.getValidHostFaceEntries(slot, prevSlots, { excludeOccupied: true, excludeSlotId: slot.id });
                if (entries.length > 0) {
                    const currentEntry = entries.find((entry) => entry.slotId === slot.hostSlotId && entry.faceId === slot.hostFaceId) || entries[0];
                    const currentLabel = currentEntry?.label || 'Face';

                    if (entries.length > 1) {
                        const cycleBtn = document.createElement('button');
                        cycleBtn.type = 'button';
                        cycleBtn.className = 'card-cycle-btn';
                        cycleBtn.dataset.cycleSlotId = String(slot.id);
                        cycleBtn.textContent = `Cycle: ${currentLabel}`;
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

    updateSliderFill(input) {
        const min = Number(input.min);
        const max = Number(input.max);
        const value = Number(input.value);
        const ratio = max === min ? 0 : (value - min) / (max - min);

        input.style.setProperty('--slider-fill', `${Math.max(0, Math.min(1, ratio)) * 100}%`);
    }

    refreshSliderBadges() {
        this.primitiveCardsListEl.querySelectorAll('.slider-input').forEach((input) => {
            this.updateSliderFill(input);
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

        this.compositeSlots.forEach((slot, idx) => {
            let entry = null;
            if (idx > 0) {
                const prevSlots = this.compositeSlots.slice(0, idx);
                entry = this.ensureSlotHostBinding(slot, prevSlots);
                if (entry) {
                    this.snapSlotDimensions(slot);
                }
            }

            const def = this.createSlotDefinition(slot);

            if (idx > 0) {
                if (entry) {
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
                    const hostFaceNormal = this.resolveFaceNormal(entry.faceDef, entry.slot.params)
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ);
                    const hostFaceU = this.resolveFaceUAxis(entry.faceDef, entry.slot.params)
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ)
                        .normalize();

                    this.applySlotTransform(def.group, slot, hostFaceCenter, hostFaceNormal, hostFaceU);

                    const guestFaceDef = this.getGuestAttachFaceDef(slot, this.resolveFaceNormal(entry.faceDef, entry.slot.params));
                    if (guestFaceDef) {
                        this.addLinkages(
                            { slotId: entry.slotId, dims: this.resolveFaceDims(entry.faceDef, entry.slot.params) },
                            { slotId: slot.id, dims: this.resolveFaceDims(guestFaceDef, slot.params) }
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
                id: `s${slot.id}_${pt.id}`,
                label: pt.label,
                position: pt.position.clone().applyQuaternion(qRot).add(vPos),
            }));
            allPoints = allPoints.concat(worldPoints);
            maxBoundsRadius = Math.max(maxBoundsRadius, def.boundsRadius + def.group.position.length());
        });

        this.pointDefinitions = this.mergeCoincidentBasePoints(allPoints);
        this.rebuildConstructions();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.updatePanelCopy();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
        if (fitCamera) {
            this.fitCameraToObject(this.compositeGroup);
        }
    }

    applySlotTransform(slotGroup, slot, hostFaceCenter, hostFaceNormal, hostFaceUWorld = null) {
        const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal);
        if (!guestFaceDef) return;

        const targetGuestNormal = hostFaceNormal.clone().negate();
        
        // Apply guest's orientation rotation to its face normal (which is in standard space)
        const guestOrientQ = this.getOrientationQuaternion(slot.primitive, slot.orientation);
        const guestFaceNormal = this.resolveFaceNormal(guestFaceDef, slot.params).applyQuaternion(guestOrientQ);
        
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

        // After normal alignment, align in-plane axis too (prevents left/right face mismatch).
        if (hostFaceUWorld && guestFaceDef.uAxis) {
            const guestUAfterNormalAlign = this.resolveFaceUAxis(guestFaceDef, slot.params)
                .applyQuaternion(guestOrientQ)
                .applyQuaternion(Q);

            const projectToPlane = (vec, normal) => vec.clone().sub(normal.clone().multiplyScalar(vec.dot(normal)));
            const guestUProj = projectToPlane(guestUAfterNormalAlign, targetGuestNormal).normalize();
            const hostUProj = projectToPlane(hostFaceUWorld, targetGuestNormal).normalize();

            if (guestUProj.lengthSq() > 1e-8 && hostUProj.lengthSq() > 1e-8) {
                const cross = new THREE.Vector3().crossVectors(guestUProj, hostUProj);
                const sin = targetGuestNormal.dot(cross);
                const cos = THREE.MathUtils.clamp(guestUProj.dot(hostUProj), -1, 1);
                const twistAngle = Math.atan2(sin, cos);
                const twist = new THREE.Quaternion().setFromAxisAngle(targetGuestNormal, twistAngle);
                Q.premultiply(twist);
            }
        }

        slotGroup.quaternion.copy(Q);
        const guestLocalCenter = guestFaceDef.center(slot.params)
            .clone()
            .applyQuaternion(guestOrientQ);  // Apply orientation first
        const guestRotatedCenter = guestLocalCenter.applyQuaternion(Q);
        slotGroup.position.copy(hostFaceCenter).sub(guestRotatedCenter);
    }

    addLinkages(hostInfo, guestInfo) {
        const hostSlot = this.compositeSlots.find((s) => s.id === hostInfo.slotId);
        const guestSlot = this.compositeSlots.find((s) => s.id === guestInfo.slotId);
        if (!hostSlot || !guestSlot) return;

        const len = Math.min(hostInfo.dims.length, guestInfo.dims.length);
        for (let i = 0; i < len; i++) {
            const hostDim = hostInfo.dims[i];
            const guestDim = guestInfo.dims[i];
            const hostHasParam = hostDim && Object.prototype.hasOwnProperty.call(hostSlot.params, hostDim);
            const guestHasParam = guestDim && Object.prototype.hasOwnProperty.call(guestSlot.params, guestDim);
            if (hostHasParam && guestHasParam) {
                this.slotLinkages.push({ fromSlotId: hostInfo.slotId, fromParam: hostDim, toSlotId: guestInfo.slotId, toParam: guestDim });
                this.slotLinkages.push({ fromSlotId: guestInfo.slotId, fromParam: guestDim, toSlotId: hostInfo.slotId, toParam: hostDim });
            }
        }
    }
    
    getUniqueDisplayLabel(preferredLabel, usedLabels) {
        if (preferredLabel && !usedLabels.has(preferredLabel)) {
            usedLabels.add(preferredLabel);
            return preferredLabel;
        }
        
        const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        for (const letter of alphabet) {
            if (!usedLabels.has(letter)) {
                usedLabels.add(letter);
                return letter;
            }
        }
        
        let index = 1;
        while (usedLabels.has(`P${index}`)) {
            index += 1;
        }
        const fallback = `P${index}`;
        usedLabels.add(fallback);
        return fallback;
    }
    
    mergeCoincidentBasePoints(rawPoints, threshold = 0.02) {
        const merged = [];
        const usedLabels = new Set();
        let mergedIndex = 1;
        
        rawPoints.forEach((point) => {
            const existing = merged.find((candidate) => candidate.position.distanceTo(point.position) <= threshold);
            if (existing) {
                existing.sourceIds.push(point.id);
                return;
            }
            
            const label = this.getUniqueDisplayLabel(point.label, usedLabels);
            merged.push({
                id: `p${mergedIndex}`,
                label,
                description: point.description,
                position: point.position.clone(),
                sourceIds: [point.id]
            });
            mergedIndex += 1;
        });
        
        return merged;
    }
    
    formatPointSequence(pointIds) {
        return pointIds
            .map((pointId) => this.getPointById(pointId)?.label || pointId)
            .join('');
    }

    normalizePointPairIds(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 2) {
            return null;
        }

        return [...pointIds].sort((left, right) => String(left).localeCompare(String(right)));
    }

    getPrimitiveEdgeSet(primitiveKey) {
        const edgePairsByPrimitive = {
            cuboid: [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['E', 'F'], ['F', 'G'], ['G', 'H'], ['H', 'E'],
                ['A', 'E'], ['B', 'F'], ['C', 'G'], ['D', 'H']
            ],
            'right-triangle-prism': [
                ['A', 'B'], ['B', 'C'], ['C', 'A'],
                ['D', 'E'], ['E', 'F'], ['F', 'D'],
                ['A', 'D'], ['B', 'E'], ['C', 'F']
            ],
            tetrahedron: [
                ['A', 'B'], ['B', 'C'], ['C', 'A'],
                ['A', 'D'], ['B', 'D'], ['C', 'D']
            ],
            'trapezium-prism': [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['E', 'F'], ['F', 'G'], ['G', 'H'], ['H', 'E'],
                ['A', 'E'], ['B', 'F'], ['C', 'G'], ['D', 'H']
            ],
            sphere: [
                ['A', 'C'], ['A', 'D'], ['A', 'E'], ['A', 'F'],
                ['B', 'C'], ['B', 'D'], ['B', 'E'], ['B', 'F'],
                ['C', 'D'], ['D', 'E'], ['E', 'F'], ['F', 'C'],
                ['A', 'O'], ['B', 'O'], ['C', 'O'], ['D', 'O'], ['E', 'O'], ['F', 'O']
            ],
            hemisphere: [
                ['A', 'B'], ['A', 'C'], ['A', 'D'], ['A', 'E'],
                ['B', 'C'], ['C', 'D'], ['D', 'E'], ['E', 'B'],
                ['B', 'F'], ['C', 'F'], ['D', 'F'], ['E', 'F'], ['A', 'F']
            ],
            cylinder: [
                ['A', 'B'], ['A', 'C'], ['A', 'D'], ['B', 'E'], ['B', 'F'],
                ['C', 'D'], ['E', 'F'], ['C', 'E'], ['D', 'F'],
                ['A', 'O'], ['B', 'O'], ['C', 'O'], ['D', 'O'], ['E', 'O'], ['F', 'O']
            ],
            cone: [
                ['A', 'B'], ['A', 'C'], ['A', 'D'], ['A', 'E'], ['A', 'F'],
                ['B', 'C'], ['B', 'D'], ['B', 'E'], ['B', 'F'],
                ['C', 'D'], ['D', 'E'], ['E', 'F'], ['F', 'C']
            ],
            'rectangular-pyramid': [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['A', 'E'], ['B', 'E'], ['C', 'E'], ['D', 'E']
            ]
        };

        const pairs = edgePairsByPrimitive[primitiveKey] || [];
        return new Set(pairs.map((pair) => pair.slice().sort().join('|')));
    }

    getPrimitiveFacePointSets(primitiveKey) {
        const facePointSetsByPrimitive = {
            cuboid: [
                ['A', 'B', 'C', 'D'],
                ['E', 'F', 'G', 'H'],
                ['A', 'B', 'F', 'E'],
                ['B', 'C', 'G', 'F'],
                ['C', 'D', 'H', 'G'],
                ['D', 'A', 'E', 'H']
            ],
            'right-triangle-prism': [
                ['A', 'B', 'C'],
                ['D', 'E', 'F'],
                ['A', 'B', 'E', 'D'],
                ['B', 'C', 'F', 'E'],
                ['C', 'A', 'D', 'F']
            ],
            tetrahedron: [
                ['A', 'B', 'C'],
                ['A', 'B', 'D'],
                ['B', 'C', 'D'],
                ['C', 'A', 'D']
            ],
            'trapezium-prism': [
                ['A', 'B', 'C', 'D'],
                ['E', 'F', 'G', 'H'],
                ['A', 'B', 'F', 'E'],
                ['B', 'C', 'G', 'F'],
                ['C', 'D', 'H', 'G'],
                ['D', 'A', 'E', 'H']
            ],
            'rectangular-pyramid': [
                ['A', 'B', 'C', 'D'],
                ['A', 'B', 'E'],
                ['B', 'C', 'E'],
                ['C', 'D', 'E'],
                ['D', 'A', 'E']
            ],
            sphere: [
                ['A', 'B', 'C', 'D', 'E', 'F']
            ],
            hemisphere: [
                ['A', 'B', 'C', 'D', 'E'],
                ['B', 'C', 'D', 'E', 'F']
            ],
            cylinder: [
                ['A', 'C', 'D'],
                ['B', 'E', 'F'],
                ['C', 'D', 'E', 'F']
            ],
            cone: [
                ['A', 'C', 'D', 'E', 'F'],
                ['B', 'C', 'D', 'E', 'F']
            ]
        };

        return (facePointSetsByPrimitive[primitiveKey] || []).map((face) => new Set(face));
    }

    parsePointSourceId(sourceId) {
        const match = /^s(\d+)_(.+)$/.exec(String(sourceId));
        if (!match) return null;
        return { slotId: match[1], localId: match[2] };
    }

    hasPrimitiveEdgeBetween(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return false;
        }

        const pointA = this.getPointById(normalized[0]);
        const pointB = this.getPointById(normalized[1]);
        if (!pointA?.sourceIds?.length || !pointB?.sourceIds?.length) {
            return false;
        }

        const slotPrimitiveMap = new Map(this.compositeSlots.map((slot) => [String(slot.id), slot.primitive]));
        const sourceA = pointA.sourceIds.map((sourceId) => this.parsePointSourceId(sourceId)).filter(Boolean);
        const sourceB = pointB.sourceIds.map((sourceId) => this.parsePointSourceId(sourceId)).filter(Boolean);

        for (const left of sourceA) {
            for (const right of sourceB) {
                if (left.slotId !== right.slotId || left.localId === right.localId) {
                    continue;
                }

                const primitiveKey = slotPrimitiveMap.get(left.slotId);
                if (!primitiveKey) {
                    continue;
                }

                const edgeKey = [left.localId, right.localId].sort().join('|');
                if (this.getPrimitiveEdgeSet(primitiveKey).has(edgeKey)) {
                    return true;
                }
            }
        }

        return false;
    }

    hasPrimitiveFaceBetween(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return false;
        }

        const pointA = this.getPointById(normalized[0]);
        const pointB = this.getPointById(normalized[1]);
        if (!pointA?.sourceIds?.length || !pointB?.sourceIds?.length) {
            return false;
        }

        const slotPrimitiveMap = new Map(this.compositeSlots.map((slot) => [String(slot.id), slot.primitive]));
        const sourceA = pointA.sourceIds.map((sourceId) => this.parsePointSourceId(sourceId)).filter(Boolean);
        const sourceB = pointB.sourceIds.map((sourceId) => this.parsePointSourceId(sourceId)).filter(Boolean);

        for (const left of sourceA) {
            for (const right of sourceB) {
                if (left.slotId !== right.slotId || left.localId === right.localId) {
                    continue;
                }

                const primitiveKey = slotPrimitiveMap.get(left.slotId);
                if (!primitiveKey) {
                    continue;
                }

                const faceSets = this.getPrimitiveFacePointSets(primitiveKey);
                const isSameFacePair = faceSets.some((faceSet) => faceSet.has(left.localId) && faceSet.has(right.localId));
                if (isSameFacePair) {
                    return true;
                }
            }
        }

        return false;
    }

    hasSceneSegmentBetween(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return false;
        }

        const matchesPair = (pairCandidate) => {
            const pair = this.normalizePointPairIds(pairCandidate);
            return !!pair && pair[0] === normalized[0] && pair[1] === normalized[1];
        };

        return this.sceneObjects.some((entry) => {
            const definition = entry.definition;
            if (!definition || !Array.isArray(definition.pointIds)) {
                return false;
            }

            if (definition.kind === 'segment' && definition.pointIds.length === 2) {
                return matchesPair(definition.pointIds);
            }

            if (definition.kind === 'triangle' && definition.pointIds.length === 3) {
                const ids = definition.pointIds;
                return matchesPair([ids[0], ids[1]])
                    || matchesPair([ids[1], ids[2]])
                    || matchesPair([ids[2], ids[0]]);
            }

            if (definition.kind === 'plane' && definition.pointIds.length === 4) {
                const ids = definition.pointIds;
                return matchesPair([ids[0], ids[1]])
                    || matchesPair([ids[1], ids[2]])
                    || matchesPair([ids[2], ids[3]])
                    || matchesPair([ids[3], ids[0]]);
            }

            return false;
        });
    }

    canAttachLabelToPointPair(pointIds) {
        return this.hasPrimitiveEdgeBetween(pointIds) || this.hasPrimitiveFaceBetween(pointIds) || this.hasSceneSegmentBetween(pointIds);
    }

    hasSegmentLikeConnection(pointIds) {
        return this.hasPrimitiveEdgeBetween(pointIds) || this.hasSceneSegmentBetween(pointIds);
    }

    canAttachAngleFromOrderedPoints(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 3) {
            return false;
        }

        if (new Set(pointIds).size !== 3) {
            return false;
        }

        const leftLeg = [pointIds[0], pointIds[1]];
        const rightLeg = [pointIds[1], pointIds[2]];
        return this.hasSegmentLikeConnection(leftLeg) && this.hasSegmentLikeConnection(rightLeg);
    }

    findEdgeLabelObject(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return null;
        }

        return this.sceneObjects.find((entry) => {
            const definition = entry.definition;
            if (!definition || (definition.kind !== 'edge-label' && definition.kind !== 'length-label') || !Array.isArray(definition.pointIds) || definition.pointIds.length !== 2) {
                return false;
            }

            const pair = this.normalizePointPairIds(definition.pointIds);
            return !!pair && pair[0] === normalized[0] && pair[1] === normalized[1];
        }) || null;
    }

    hasMidpointForPair(pointIds) {
        const signature = this.makeMidpointSignature(pointIds);
        if (!signature) {
            return false;
        }

        return this.sceneObjects.some((entry) => entry.definition?.kind === 'midpoint-point' && entry.definition.signature === signature);
    }

    getMidpointSignatureForPointId(pointId) {
        const point = this.getPointById(pointId);
        if (!point || !point.isDerived) {
            return null;
        }

        if (typeof point.signature === 'string' && point.signature.startsWith('midpoint|')) {
            return point.signature;
        }

        if (typeof point.id === 'string' && point.id.startsWith('derived-midpoint|')) {
            return point.id.replace('derived-', '');
        }

        return null;
    }

    findMidpointObjectByPointId(pointId) {
        const signature = this.getMidpointSignatureForPointId(pointId);
        if (!signature) {
            return null;
        }

        return this.sceneObjects.find((entry) => entry.definition?.kind === 'midpoint-point' && entry.definition.signature === signature) || null;
    }

    removeEdgeLabelsForPointPair(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return;
        }

        const survivors = [];
        this.sceneObjects.forEach((entry) => {
            const definition = entry.definition;
            const isEdgeLabel = definition && (definition.kind === 'edge-label' || definition.kind === 'length-label') && Array.isArray(definition.pointIds) && definition.pointIds.length === 2;
            if (!isEdgeLabel) {
                survivors.push(entry);
                return;
            }

            const pair = this.normalizePointPairIds(definition.pointIds);
            const isTargetPair = !!pair && pair[0] === normalized[0] && pair[1] === normalized[1];
            if (!isTargetPair) {
                survivors.push(entry);
                return;
            }

            this.scene.remove(entry.object3D);
            this.disposeObject3D(entry.object3D);
        });

        this.sceneObjects = survivors;
    }

    removeMidpointPointsForPair(pointIds, options = {}) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return;
        }

        const onlyIfDisconnected = options.onlyIfDisconnected === true;
        if (onlyIfDisconnected && this.canAttachLabelToPointPair(normalized)) {
            return;
        }

        const signature = this.makeMidpointSignature(normalized);
        const survivors = [];
        this.sceneObjects.forEach((entry) => {
            if (entry.definition?.kind !== 'midpoint-point' || entry.definition.signature !== signature) {
                survivors.push(entry);
                return;
            }

            this.scene.remove(entry.object3D);
            this.disposeObject3D(entry.object3D);
        });

        this.sceneObjects = survivors;
    }

    getTriangleEdgePairs(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 3) {
            return [];
        }

        return [
            [pointIds[0], pointIds[1]],
            [pointIds[1], pointIds[2]],
            [pointIds[2], pointIds[0]]
        ];
    }

    getPlaneEdgePairs(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 4) {
            return [];
        }

        return [
            [pointIds[0], pointIds[1]],
            [pointIds[1], pointIds[2]],
            [pointIds[2], pointIds[3]],
            [pointIds[3], pointIds[0]]
        ];
    }

    ensureHiddenSupportSegment(pointIds, reasonLabel = 'Support edge', ownerId = null) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return;
        }

        if (this.canAttachLabelToPointPair(normalized)) {
            return;
        }

        this.addSceneObject({
            type: 'segment',
            name: `Ghost ${this.formatPointSequence(normalized)}`,
            subtitle: reasonLabel,
            object3D: new THREE.Group(),
            definition: {
                kind: 'segment',
                pointIds: normalized,
                hidden: true,
                supportOwnerId: ownerId
            }
        });
    }

    ensureHiddenSupportSegmentsForPairs(pairs, reasonLabel = 'Support edge', ownerId = null) {
        if (!Array.isArray(pairs) || pairs.length === 0) {
            return;
        }

        pairs.forEach((pair) => {
            this.ensureHiddenSupportSegment(pair, reasonLabel, ownerId);
        });
    }

    removeHiddenSupportSegmentsForOwner(ownerId) {
        if (!Number.isFinite(ownerId)) {
            return;
        }

        const survivors = [];
        const removedPairs = [];
        this.sceneObjects.forEach((entry) => {
            const definition = entry.definition;
            const isOwnedHiddenSupportSegment = !!definition
                && definition.kind === 'segment'
                && definition.hidden === true
                && definition.supportOwnerId === ownerId;

            if (!isOwnedHiddenSupportSegment) {
                survivors.push(entry);
                return;
            }

            const pair = this.normalizePointPairIds(definition.pointIds || []);
            if (pair) {
                removedPairs.push(pair);
            }

            this.scene.remove(entry.object3D);
            this.disposeObject3D(entry.object3D);
        });

        this.sceneObjects = survivors;

        removedPairs.forEach((pair) => {
            this.removeEdgeLabelsForPointPair(pair);
        });
    }

    normalizePointLabelInput(rawLabel) {
        if (typeof rawLabel !== 'string') {
            return null;
        }

        const trimmed = rawLabel.trim().toUpperCase();
        if (!/^[A-Z](?:\d)?$/.test(trimmed)) {
            return null;
        }

        return trimmed;
    }

    isPointLabelAvailable(label, excludePointId = null) {
        return !this.getAllPoints().some((point) => point.id !== excludePointId && point.label === label);
    }

    makeDerivedSignature(firstSegment, secondSegment, intersectionPoint) {
        const firstKey = [...firstSegment.pointIds].sort().join('-');
        const secondKey = [...secondSegment.pointIds].sort().join('-');
        const [leftKey, rightKey] = [firstKey, secondKey].sort();
        const x = intersectionPoint.x.toFixed(3);
        const y = intersectionPoint.y.toFixed(3);
        const z = intersectionPoint.z.toFixed(3);
        return `${leftKey}|${rightKey}|${x}|${y}|${z}`;
    }

    makeMidpointSignature(pointIds) {
        const normalized = this.normalizePointPairIds(pointIds);
        return normalized ? `midpoint|${normalized[0]}|${normalized[1]}` : null;
    }

    changeSelectedPointLabel() {
        const pointId = this.selectedPoints[0];
        const point = this.getPointById(pointId);
        if (!point) {
            return;
        }

        const nextLabelRaw = window.prompt(`Change label for point ${point.label}`, point.label);
        if (nextLabelRaw == null) {
            return;
        }

        const nextLabel = this.normalizePointLabelInput(nextLabelRaw);
        if (!nextLabel) {
            window.alert('Use label format A or A1 (single letter, optional single digit).');
            return;
        }

        if (!this.isPointLabelAvailable(nextLabel, point.id)) {
            window.alert(`Label ${nextLabel} is already in use.`);
            return;
        }

        if (point.isDerived) {
            point.label = nextLabel;
            if (point.signature) {
                this.derivedLabelOverrides.set(point.signature, nextLabel);
            }
        } else {
            point.label = nextLabel;
            this.refreshDerivedPoints();
        }

        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    updatePanelCopy() {
        if (this.compositeSlots.length === 0) {
            this.primitiveChip.textContent = '-';
            this.orientationChip.textContent = '-';
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
        let intrinsicRightAngleEdgePairs = [];

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
            intrinsicRightAngleEdgePairs = [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['E', 'F'], ['F', 'G'], ['G', 'H'], ['H', 'E'],
                ['A', 'E'], ['B', 'F'], ['C', 'G'], ['D', 'H']
            ];
            boundsRadius = Math.max(width, depth, height) * 1.15;
        } else if (primitiveKey === 'right-triangle-prism') {
            const { legA, legB, length, triangleMode } = params;
            const zFront = length / 2;
            const zBack = -length / 2;

            const mode = normalizeTriangularPrismMode(triangleMode);
            const [posA, posB, posC] = getTriangularPrismProfilePoints(params, zFront);

            const posD = new THREE.Vector3(posA.x, posA.y, zBack);
            const posE = new THREE.Vector3(posB.x, posB.y, zBack);
            const posF = new THREE.Vector3(posC.x, posC.y, zBack);

            const vertices = [
                posA.x, posA.y, posA.z,
                posB.x, posB.y, posB.z,
                posC.x, posC.y, posC.z,
                posD.x, posD.y, posD.z,
                posE.x, posE.y, posE.z,
                posF.x, posF.y, posF.z
            ];

            const indices = [
                0, 1, 2,
                3, 5, 4,
                0, 1, 4,
                0, 4, 3,
                0, 3, 5,
                0, 5, 2,
                1, 2, 5,
                1, 5, 4
            ];

            geometry = new THREE.BufferGeometry();
            geometry.setAttribute('position', new THREE.Float32BufferAttribute(vertices, 3));
            geometry.setIndex(indices);
            geometry.computeVertexNormals();

            const modeDesc = mode === 'isosceles' ? 'isosceles' : mode === 'right-above-B' ? 'right angle at B' : 'right angle at A';
            points = [
                { id: 'A', label: 'A', description: `front base A (${modeDesc})`, position: posA },
                { id: 'B', label: 'B', description: `front base B (${modeDesc})`, position: posB },
                { id: 'C', label: 'C', description: `front apex C (${modeDesc})`, position: posC },
                { id: 'D', label: 'D', description: `back base A (${modeDesc})`, position: posD },
                { id: 'E', label: 'E', description: `back base B (${modeDesc})`, position: posE },
                { id: 'F', label: 'F', description: `back apex C (${modeDesc})`, position: posF }
            ];
            intrinsicRightAngleEdgePairs = [
                ['A', 'B'], ['B', 'C'], ['C', 'A'],
                ['D', 'E'], ['E', 'F'], ['F', 'D'],
                ['A', 'D'], ['B', 'E'], ['C', 'F']
            ];

            boundsRadius = Math.max(legA, legB, length) * 1.2;
        } else if (primitiveKey === 'tetrahedron') {
            const { height, baseTriangleMode, apexPosition } = params;
            const yBase = -height / 2;
            const yApex = height / 2;
            const [baseA, baseB, baseC] = getTetrahedronBasePoints(params, yBase);
            const apexTargetKey = apexPosition || 'A';
            const apexTargets = { A: baseA, B: baseB, C: baseC };
            const apexAnchor = apexTargets[apexTargetKey] || baseA;
            const apex = new THREE.Vector3(apexAnchor.x, yApex, apexAnchor.z);

            const vertices = [
                baseA.x, baseA.y, baseA.z,
                baseB.x, baseB.y, baseB.z,
                baseC.x, baseC.y, baseC.z,
                apex.x, apex.y, apex.z
            ];

            const indices = [
                0, 2, 1,
                0, 1, 3,
                1, 2, 3,
                2, 0, 3
            ];

            geometry = new THREE.BufferGeometry();
            geometry.setAttribute('position', new THREE.Float32BufferAttribute(vertices, 3));
            geometry.setIndex(indices);
            geometry.computeVertexNormals();

            const baseMode = normalizeTetrahedronBaseMode(baseTriangleMode);
            const modeDesc = baseMode === 'right-angled' ? 'right-angled base' : 'isosceles base';
            const apexDesc = `apex ${this.tetrahedronApexPositions.find((opt) => opt.value === apexTargetKey)?.label?.toLowerCase() || 'above A'}`;
            points = [
                { id: 'A', label: 'A', description: `base vertex A (${modeDesc})`, position: baseA },
                { id: 'B', label: 'B', description: `base vertex B (${modeDesc})`, position: baseB },
                { id: 'C', label: 'C', description: `base vertex C (${modeDesc})`, position: baseC },
                { id: 'D', label: 'D', description: apexDesc, position: apex }
            ];
            intrinsicRightAngleEdgePairs = [
                ['A', 'B'], ['B', 'C'], ['C', 'A'],
                ['A', 'D'], ['B', 'D'], ['C', 'D']
            ];

            boundsRadius = Math.max(params.base, params.triangleHeight, height) * 1.2;
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

            intrinsicRightAngleEdgePairs = [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['E', 'F'], ['F', 'G'], ['G', 'H'], ['H', 'E'],
                ['A', 'E'], ['B', 'F'], ['C', 'G'], ['D', 'H']
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
            const yBase = -height / 2;
            const yApex = height / 2;
            const baseA = new THREE.Vector3(-length / 2, yBase, width / 2);
            const baseB = new THREE.Vector3(length / 2, yBase, width / 2);
            const baseC = new THREE.Vector3(length / 2, yBase, -width / 2);
            const baseD = new THREE.Vector3(-length / 2, yBase, -width / 2);
            const apexTargets = {
                center: new THREE.Vector3(0, yApex, 0),
                A: new THREE.Vector3(baseA.x, yApex, baseA.z),
                B: new THREE.Vector3(baseB.x, yApex, baseB.z),
                C: new THREE.Vector3(baseC.x, yApex, baseC.z),
                D: new THREE.Vector3(baseD.x, yApex, baseD.z)
            };
            const selectedApexPosition = params.apexPosition || 'center';
            const apex = (apexTargets[selectedApexPosition] || apexTargets.center).clone();

            const localPoints = [baseA.clone(), baseB.clone(), baseC.clone(), baseD.clone(), apex.clone()];
            if (slot.orientation === 'apex-down') {
                localPoints.forEach((pt) => {
                    pt.x *= -1;
                    pt.y *= -1;
                });
            }

            const [A, B, C, D, E] = localPoints;
            const flatVertices = [
                A.x, A.y, A.z,
                B.x, B.y, B.z,
                C.x, C.y, C.z,
                D.x, D.y, D.z,
                E.x, E.y, E.z
            ];

            const indices = [
                0, 2, 1,
                0, 3, 2,
                0, 1, 4,
                1, 2, 4,
                2, 3, 4,
                3, 0, 4
            ];

            geometry = new THREE.BufferGeometry();
            geometry.setAttribute('position', new THREE.Float32BufferAttribute(flatVertices, 3));
            geometry.setIndex(indices);
            geometry.computeVertexNormals();

            const baseCentre = new THREE.Vector3(0, yBase, 0);
            if (slot.orientation === 'apex-down') {
                baseCentre.x *= -1;
                baseCentre.y *= -1;
            }

            const baseDescPrefix = slot.orientation === 'apex-down' ? 'top' : 'base';
            points = [
                { id: 'A', label: 'A', description: `${baseDescPrefix} front left`, position: A },
                { id: 'B', label: 'B', description: `${baseDescPrefix} front right`, position: B },
                { id: 'C', label: 'C', description: `${baseDescPrefix} back right`, position: C },
                { id: 'D', label: 'D', description: `${baseDescPrefix} back left`, position: D },
                { id: 'E', label: 'E', description: 'apex', position: E },
                { id: 'O', label: 'O', description: 'base centre', position: baseCentre }
            ];
            intrinsicRightAngleEdgePairs = [
                ['A', 'B'], ['B', 'C'], ['C', 'D'], ['D', 'A'],
                ['A', 'E'], ['B', 'E'], ['C', 'E'], ['D', 'E']
            ];
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

        if (intrinsicRightAngleEdgePairs.length > 0) {
            const pointMap = new Map(points.map((point) => [point.id, point.position]));
            const intrinsicRightAngleTriples = this.collectRightAngleTriples(pointMap, intrinsicRightAngleEdgePairs);
            const sharedMarkerSizeByVertex = new Map();
            const primitiveCenterLocal = new THREE.Vector3();
            if (points.length > 0) {
                points.forEach((point) => primitiveCenterLocal.add(point.position));
                primitiveCenterLocal.multiplyScalar(1 / points.length);
            }

            intrinsicRightAngleTriples.forEach(([vertexId, armId1, armId2]) => {
                const vertex = pointMap.get(vertexId);
                const armPoint1 = pointMap.get(armId1);
                const armPoint2 = pointMap.get(armId2);
                if (!vertex || !armPoint1 || !armPoint2) return;

                const len1 = armPoint1.distanceTo(vertex);
                const len2 = armPoint2.distanceTo(vertex);
                if (len1 < 1e-6 || len2 < 1e-6) return;

                const candidateSize = THREE.MathUtils.clamp(Math.min(len1, len2) * 0.16, 0.15, 0.9);
                const existingSize = sharedMarkerSizeByVertex.get(vertexId);
                if (existingSize == null || candidateSize < existingSize) {
                    sharedMarkerSizeByVertex.set(vertexId, candidateSize);
                }
            });

            intrinsicRightAngleTriples.forEach(([vertexId, armId1, armId2]) => {
                const vertex = pointMap.get(vertexId);
                const armPoint1 = pointMap.get(armId1);
                const armPoint2 = pointMap.get(armId2);
                if (!vertex || !armPoint1 || !armPoint2) return;

                const arm1 = armPoint1.clone().sub(vertex);
                const arm2 = armPoint2.clone().sub(vertex);
                const faceNormalLocal = new THREE.Vector3().crossVectors(arm1, arm2);
                if (faceNormalLocal.lengthSq() < 1e-10) return;
                faceNormalLocal.normalize();

                const facePointLocal = vertex.clone().add(armPoint1).add(armPoint2).multiplyScalar(1 / 3);
                const outwardRef = facePointLocal.clone().sub(primitiveCenterLocal);
                if (faceNormalLocal.dot(outwardRef) < 0) {
                    faceNormalLocal.multiplyScalar(-1);
                }

                const marker = this.createRightAngleMarker(
                    vertex,
                    armPoint1,
                    armPoint2,
                    sharedMarkerSizeByVertex.get(vertexId),
                    true
                );
                if (marker) {
                    marker.userData.isIntrinsicRightAngleMarker = true;
                    marker.userData.faceNormalLocal = faceNormalLocal;
                    marker.userData.facePointLocal = facePointLocal;
                    group.add(marker);
                }
            });
        }

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

    updateIntrinsicRightAngleMarkerVisibility() {
        if (!this.scene || !this.camera) return;

        this.scene.updateMatrixWorld(true);

        const facePointWorld = new THREE.Vector3();
        const faceNormalWorld = new THREE.Vector3();
        const toCamera = new THREE.Vector3();
        const hostWorldQ = new THREE.Quaternion();

        this.scene.traverse((obj) => {
            if (!obj.userData?.isIntrinsicRightAngleMarker) return;

            const facePointLocal = obj.userData.facePointLocal;
            const faceNormalLocal = obj.userData.faceNormalLocal;
            const hostGroup = obj.parent;
            if (!facePointLocal || !faceNormalLocal || !hostGroup) {
                obj.visible = true;
                return;
            }

            facePointWorld.copy(facePointLocal).applyMatrix4(hostGroup.matrixWorld);
            hostGroup.getWorldQuaternion(hostWorldQ);
            faceNormalWorld.copy(faceNormalLocal).applyQuaternion(hostWorldQ).normalize();
            toCamera.copy(this.camera.position).sub(facePointWorld);
            obj.visible = faceNormalWorld.dot(toCamera) >= 0;
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
            pill.textContent = this.getPointById(pointId)?.label || pointId;
            this.selectionSummaryEl.appendChild(pill);

            if (index < this.selectedPoints.length - 1) {
                const arrow = document.createElement('span');
                arrow.className = 'selection-arrow';
                arrow.textContent = '->';
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
        const baseActions = [...(this.actionsByCount[this.selectedPoints.length] || [])];

        if (this.selectedPoints.length === 1) {
            const midpointObject = this.findMidpointObjectByPointId(this.selectedPoints[0]);
            if (midpointObject) {
                baseActions.push({ key: 'delete-midpoint', label: 'Delete Midpoint' });
            }
        }

        if (this.selectedPoints.length === 2 && this.canAttachLabelToPointPair(this.selectedPoints)) {
            const hasExistingLabel = !!this.findEdgeLabelObject(this.selectedPoints);
            baseActions.push({ key: 'edge-label', label: hasExistingLabel ? 'Change Label' : 'Add Label' });

            if (!this.hasMidpointForPair(this.selectedPoints)) {
                baseActions.push({ key: 'add-midpoint', label: 'Add Midpoint' });
            }
        }

        if (this.selectedPoints.length === 3) {
            return baseActions.filter((action) => {
                if (action.key !== 'angle') {
                    return true;
                }
                return this.canAttachAngleFromOrderedPoints(this.selectedPoints);
            });
        }

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
        const priorPointMap = new Map(this.getAllPoints().map((point) => [point.id, point.position.clone()]));
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

        const derived = [];
        const activeDerivedSignatures = new Set();

        const midpointDefinitions = this.sceneObjects
            .map((entry) => entry.definition)
            .filter((definition) => definition?.kind === 'midpoint-point' && Array.isArray(definition.pointIds) && definition.pointIds.length === 2);
        const seenMidpointSignatures = new Set();

        midpointDefinitions.forEach((definition) => {
            if (!this.canAttachLabelToPointPair(definition.pointIds)) {
                return;
            }

            const signature = definition.signature || this.makeMidpointSignature(definition.pointIds);
            if (!signature || seenMidpointSignatures.has(signature)) {
                return;
            }

            const start = priorPointMap.get(definition.pointIds[0]);
            const end = priorPointMap.get(definition.pointIds[1]);
            if (!start || !end) {
                return;
            }

            const midpoint = start.clone().lerp(end, 0.5);
            if (existingPointNear(midpoint)) {
                return;
            }

            seenMidpointSignatures.add(signature);
            activeDerivedSignatures.add(signature);

            let label = this.derivedLabelOverrides.get(signature);
            if (label && usedLabels.has(label)) {
                this.derivedLabelOverrides.delete(signature);
                label = null;
            }

            if (!label) {
                label = nextDerivedLabel();
            } else {
                usedLabels.add(label);
            }

            const id = `derived-${signature}`;
            basePoints.set(id, midpoint.clone());
            derived.push({
                id,
                label,
                signature,
                description: 'derived midpoint',
                position: midpoint,
                isDerived: true
            });
        });

        const segments = this.getConstructionSegments(basePoints);

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

                const signature = this.makeDerivedSignature(first, second, intersection);
                activeDerivedSignatures.add(signature);

                let label = this.derivedLabelOverrides.get(signature);
                if (label && usedLabels.has(label)) {
                    this.derivedLabelOverrides.delete(signature);
                    label = null;
                }

                if (!label) {
                    label = nextDerivedLabel();
                } else {
                    usedLabels.add(label);
                }

                const id = `derived-${signature}`;
                basePoints.set(id, intersection.clone());
                derived.push({
                    id,
                    label,
                    signature,
                    description: 'derived intersection',
                    position: intersection,
                    isDerived: true
                });
            }
        }

        Array.from(this.derivedLabelOverrides.keys()).forEach((signature) => {
            if (!activeDerivedSignatures.has(signature)) {
                this.derivedLabelOverrides.delete(signature);
            }
        });

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
        
        const selectedLabels = this.selectedPoints.map((pointId) => this.getPointById(pointId)?.label || pointId);

        if (actionKey === 'change-point-label') {
            this.changeSelectedPointLabel();
            this.closePanelOnMobile();
            return;
        }

        if (actionKey === 'delete-midpoint') {
            if (this.selectedPoints.length !== 1) {
                return;
            }

            const midpointObject = this.findMidpointObjectByPointId(this.selectedPoints[0]);
            if (!midpointObject) {
                return;
            }

            this.deleteObject(midpointObject.id);
            this.clearSelection();
            this.closePanelOnMobile();
            return;
        }

        if (actionKey === 'segment') {
            const ids = [...this.selectedPoints];
            const color = this.nextConstructionColor();
            const segment = this.createSegment(vectors[0], vectors[1], color);
            this.addSceneObject({
                type: 'segment',
                name: `Segment ${this.formatPointSequence(ids)}`,
                subtitle: 'Two-point construction',
                object3D: segment,
                definition: {
                    kind: 'segment',
                    pointIds: ids,
                    color
                }
            });
        }

        if (actionKey === 'edge-label') {
            const ids = this.normalizePointPairIds([...this.selectedPoints]);
            if (!ids || !this.canAttachLabelToPointPair(ids)) {
                return;
            }

            const existingLabel = this.findEdgeLabelObject(ids);
            const currentText = existingLabel?.definition?.text || '';
            const promptText = existingLabel
                ? `Change label for ${this.formatPointSequence(ids)}`
                : `Label for ${this.formatPointSequence(ids)}`;
            const nextText = window.prompt(promptText, currentText);
            if (nextText == null || !nextText.trim()) {
                return;
            }

            const normalizedText = nextText.trim();
            const midpoint = vectors[0].clone().lerp(vectors[1], 0.5).add(new THREE.Vector3(0.2, 0.2, 0.2));
            const sprite = this.createTextSprite(normalizedText, {
                fontSize: 46,
                textColor: '#000000',
                background: '#b9f18a',
                borderColor: '#000000'
            });
            sprite.position.copy(midpoint);

            if (existingLabel) {
                existingLabel.definition = {
                    kind: 'edge-label',
                    pointIds: ids,
                    text: normalizedText
                };
                existingLabel.name = `Label ${this.formatPointSequence(ids)}`;
                existingLabel.subtitle = normalizedText;
                this.scene.remove(existingLabel.object3D);
                this.disposeObject3D(existingLabel.object3D);
                sprite.userData.sceneObjectId = existingLabel.id;
                existingLabel.object3D = sprite;
                existingLabel.object3D.visible = existingLabel.visible;
                this.scene.add(existingLabel.object3D);
                this.renderObjectsList();
                this.refreshDerivedPoints();
                this.buildPointMarkers();
                this.renderPointsList();
                this.renderSelectionSummary();
                this.renderActions();
            } else {
                this.addSceneObject({
                    type: 'label',
                    name: `Label ${this.formatPointSequence(ids)}`,
                    subtitle: normalizedText,
                    object3D: sprite,
                    definition: {
                        kind: 'edge-label',
                        pointIds: ids,
                        text: normalizedText
                    }
                });
            }
        }

        if (actionKey === 'add-midpoint') {
            const ids = this.normalizePointPairIds([...this.selectedPoints]);
            if (!ids || !this.canAttachLabelToPointPair(ids) || this.hasMidpointForPair(ids)) {
                return;
            }

            const signature = this.makeMidpointSignature(ids);
            this.addSceneObject({
                type: 'label',
                name: `Midpoint ${this.formatPointSequence(ids)}`,
                subtitle: 'Derived midpoint',
                object3D: new THREE.Group(),
                definition: {
                    kind: 'midpoint-point',
                    pointIds: ids,
                    signature
                }
            });
            const derivedId = `derived-${signature}`; 
            this.addSceneObject({
                type: 'segment',
                name: 'Ghost sub-segment A',
                subtitle: 'Midpoint sub-segment',
                object3D: new THREE.Group(),
                definition: { kind: 'segment', pointIds: [ids[0], derivedId], hidden: true }
            });
            this.addSceneObject({
                type: 'segment',
                name: 'Ghost sub-segment B',
                subtitle: 'Midpoint sub-segment',
                object3D: new THREE.Group(),
                definition: { kind: 'segment', pointIds: [derivedId, ids[1]], hidden: true }
            });
        }

        if (actionKey === 'triangle') {
            const ids = [...this.selectedPoints];
            const color = this.nextConstructionColor();
            const triangle = this.createTriangle(vectors[0], vectors[1], vectors[2], color, 0.28);
            const triangleObject = this.addSceneObject({
                type: 'triangle',
                name: `Triangle ${this.formatPointSequence(ids)}`,
                subtitle: 'Three-point section',
                object3D: triangle,
                definition: {
                    kind: 'triangle',
                    pointIds: ids,
                    color,
                    opacity: 0.28
                }
            });
            this.ensureHiddenSupportSegmentsForPairs(this.getTriangleEdgePairs(ids), 'Triangle support edge', triangleObject?.id ?? null);
        }

        if (actionKey === 'angle') {
            const ids = [...this.selectedPoints];
            if (!this.canAttachAngleFromOrderedPoints(ids)) {
                return;
            }

            const defaultAngleLabel = this.formatPointSequence(ids);
            let angleLabelInput = defaultAngleLabel;
            while (true) {
                const nextLabel = window.prompt(`Label for angle at ${selectedLabels[1]}`, angleLabelInput);
                if (nextLabel == null) {
                    return;
                }

                angleLabelInput = nextLabel.trim();
                if (angleLabelInput) {
                    break;
                }

                window.alert('Angle label cannot be blank.');
            }

            const color = this.nextConstructionColor();
            const angleGroup = this.createAngleMarker(vectors[0], vectors[1], vectors[2], angleLabelInput, color);
            this.addSceneObject({
                type: 'angle',
                name: `Angle ${angleLabelInput}`,
                subtitle: `Angle at ${selectedLabels[1]}`,
                object3D: angleGroup,
                definition: {
                    kind: 'angle',
                    pointIds: ids,
                    text: angleLabelInput,
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
            const planeObject = this.addSceneObject({
                type: 'plane',
                name: `Plane ${this.formatPointSequence(orderedIds)}`,
                subtitle: 'Four-point shading',
                object3D: plane,
                definition: {
                    kind: 'plane',
                    pointIds: orderedIds,
                    color,
                    opacity: 0.2
                }
            });
            this.ensureHiddenSupportSegmentsForPairs(this.getPlaneEdgePairs(orderedIds), 'Plane support edge', planeObject?.id ?? null);
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

        const triangleMarkerBaseSize = THREE.MathUtils.clamp(
            Math.min(a.distanceTo(b), b.distanceTo(c), c.distanceTo(a)) * 0.16,
            0.15,
            0.9
        );

        const rightAngleMarkers = [
            this.createRightAngleMarker(a, b, c, triangleMarkerBaseSize, false, color, 5, 1.5),
            this.createRightAngleMarker(b, c, a, triangleMarkerBaseSize, false, color, 5, 1.5),
            this.createRightAngleMarker(c, a, b, triangleMarkerBaseSize, false, color, 5, 1.5)
        ].filter(Boolean);

        const group = new THREE.Group();
        group.add(mesh, outline, outline2, outline3, ...rightAngleMarkers);
        return group;
    }

    collectRightAngleTriples(pointMap, edgePairs) {
        const adjacency = new Map();
        const addNeighbor = (fromId, toId) => {
            if (!adjacency.has(fromId)) adjacency.set(fromId, new Set());
            adjacency.get(fromId).add(toId);
        };

        edgePairs.forEach(([idA, idB]) => {
            if (!pointMap.has(idA) || !pointMap.has(idB)) return;
            addNeighbor(idA, idB);
            addNeighbor(idB, idA);
        });

        const rightAngleTriples = [];
        const rightAngleToleranceRadians = THREE.MathUtils.degToRad(2);

        adjacency.forEach((neighborSet, vertexId) => {
            const neighbors = Array.from(neighborSet);
            const vertex = pointMap.get(vertexId);
            if (!vertex || neighbors.length < 2) return;

            for (let i = 0; i < neighbors.length; i += 1) {
                for (let j = i + 1; j < neighbors.length; j += 1) {
                    const armPoint1 = pointMap.get(neighbors[i]);
                    const armPoint2 = pointMap.get(neighbors[j]);
                    if (!armPoint1 || !armPoint2) continue;

                    const arm1 = armPoint1.clone().sub(vertex);
                    const arm2 = armPoint2.clone().sub(vertex);
                    if (arm1.lengthSq() < 1e-8 || arm2.lengthSq() < 1e-8) continue;

                    const cosTheta = THREE.MathUtils.clamp(arm1.normalize().dot(arm2.normalize()), -1, 1);
                    const angle = Math.acos(cosTheta);
                    if (Math.abs(angle - Math.PI / 2) <= rightAngleToleranceRadians) {
                        rightAngleTriples.push([vertexId, neighbors[i], neighbors[j]]);
                    }
                }
            }
        });

        return rightAngleTriples;
    }

    createRightAngleMarker(vertex, armPoint1, armPoint2, markerSizeOverride = null, useThinLine = false, markerColor = 0x000000, markerLineWidth = 5, sizeScale = 1) {
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

        const markerSize = typeof markerSizeOverride === 'number' && Number.isFinite(markerSizeOverride)
            ? markerSizeOverride
            : THREE.MathUtils.clamp(Math.min(len1, len2) * 0.16, 0.15, 0.9);
        const finalMarkerSize = markerSize * 0.5 * sizeScale;
        const p1 = vertex.clone().add(u.clone().multiplyScalar(finalMarkerSize));
        const p2 = p1.clone().add(v.clone().multiplyScalar(finalMarkerSize));
        const p3 = vertex.clone().add(v.clone().multiplyScalar(finalMarkerSize));

        if (useThinLine) {
            const markerGeometry = new THREE.BufferGeometry().setFromPoints([p1, p2, p3]);
            const markerLine = new THREE.Line(
                markerGeometry,
                new THREE.LineBasicMaterial({ color: markerColor, transparent: false, opacity: 1 })
            );
            markerLine.renderOrder = 22;
            return markerLine;
        }

        const markerLine = this.createThickPolyline([p1, p2, p3], markerColor, markerLineWidth);
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
        const label = this.createTextSprite(`${angleText}`, {
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
        const entry = {
            id: this.nextObjectId,
            type,
            name,
            subtitle,
            object3D,
            definition,
            visible: true
        };
        this.sceneObjects.unshift(entry);
        this.nextObjectId += 1;
        this.renderObjectsList();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
        return entry;
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

        if (definition.hidden) {
            return new THREE.Group();
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

        if (definition.kind === 'length-label' || definition.kind === 'edge-label') {
            const vectors = this.getVectorsByPointIds(definition.pointIds || []);
            if (!vectors || vectors.length !== 2) return null;
            if (!this.canAttachLabelToPointPair(definition.pointIds || [])) return null;
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

        if (definition.kind === 'midpoint-point') {
            const midpointHolder = new THREE.Group();
            midpointHolder.visible = false;
            return midpointHolder;
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
            const angleText = (typeof definition.text === 'string' && definition.text.trim())
                ? definition.text.trim()
                : this.formatPointSequence(definition.pointIds || []);
            return this.createAngleMarker(vectors[0], vectors[1], vectors[2], angleText, definition.color || 0xff595e);
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
            const items = this.sceneObjects.filter((item) => {
                if (item.type !== type) return false;
                if (item.definition?.hidden) return false;
                if (item.definition?.kind === 'midpoint-point') return false;
                return true;
            });
            sec.list.innerHTML = '';
            items.forEach((item) => sec.list.appendChild(this.renderObjectItem(item)));
            sec.count.textContent = items.length > 0 ? `(${items.length})` : '';
        }
    }



    renderObjectItem(item) {
        const row = document.createElement('div');
        row.className = item.visible ? 'object-item' : 'object-item disabled';
        const itemColor = item.definition?.color != null
            ? `#${item.definition.color.toString(16).padStart(6, '0')}`
            : null;
        if (itemColor) {
            row.style.borderLeftColor = itemColor;
        }
        row.innerHTML = `
            <div class="object-name">
                <strong>${item.name}</strong>
                <span>${item.subtitle}</span>
            </div>
            <div class="object-controls">
                <button
                    type="button"
                    class="object-visibility-btn"
                    data-toggle-object-id="${item.id}"
                    aria-label="${item.visible ? 'Hide' : 'Show'} object"
                    title="Click to ${item.visible ? 'hide' : 'show'} object"
                    style="background-color: ${item.visible && itemColor ? itemColor : 'transparent'};"
                ></button>
                <button type="button" class="object-delete" data-delete-object-id="${item.id}" aria-label="Delete object" title="Delete object">X</button>
            </div>
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

    pruneOrphanedSceneObjects() {
        const validPointIds = new Set(this.getAllPoints().map((p) => p.id));
        const orphaned = [];
        const remaining = [];
        this.sceneObjects.forEach((entry) => {
            const def = entry.definition;
            if (def && Array.isArray(def.pointIds) && def.pointIds.length > 0) {
                if (def.pointIds.some((id) => !validPointIds.has(id))) {
                    orphaned.push(entry);
                    return;
                }
            }
            remaining.push(entry);
        });
        if (orphaned.length === 0) return;
        this.sceneObjects = remaining;
        orphaned.forEach((entry) => {
            this.scene.remove(entry.object3D);
            this.disposeObject3D(entry.object3D);
        });
    }

    deleteObject(objectId) {
        const index = this.sceneObjects.findIndex((entry) => entry.id === objectId);
        if (index === -1) return;

        const [item] = this.sceneObjects.splice(index, 1);
        if (item.definition?.kind === 'segment' && Array.isArray(item.definition.pointIds) && item.definition.pointIds.length === 2) {
            this.removeEdgeLabelsForPointPair(item.definition.pointIds);
            this.removeMidpointPointsForPair(item.definition.pointIds, { onlyIfDisconnected: true });
        }
        this.scene.remove(item.object3D);
        this.disposeObject3D(item.object3D);
        this.removeHiddenSupportSegmentsForOwner(item.id);
        this.refreshDerivedPoints();
        this.pruneOrphanedSceneObjects();
        this.renderObjectsList();
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

    fitCameraToObject(object3D, padding = 1.38, preferredDirection = null) {
        if (!object3D) {
            this.fitCameraToPrimitive(6);
            return;
        }

        const bounds = new THREE.Box3().setFromObject(object3D);
        if (bounds.isEmpty()) {
            this.fitCameraToPrimitive(6);
            return;
        }

        const center = bounds.getCenter(new THREE.Vector3());
        const sphere = bounds.getBoundingSphere(new THREE.Sphere());
        const radius = Math.max(0.001, sphere.radius);

        const vFov = THREE.MathUtils.degToRad(this.camera.fov);
        const hFov = 2 * Math.atan(Math.tan(vFov / 2) * this.camera.aspect);
        const fitHeightDistance = radius / Math.tan(vFov / 2);
        const fitWidthDistance = radius / Math.tan(hFov / 2);
        const distance = Math.max(6, padding * Math.max(fitHeightDistance, fitWidthDistance));

        const viewDir = preferredDirection
            ? preferredDirection.clone()
            : this.camera.position.clone().sub(this.controls.target);
        if (viewDir.lengthSq() < 1e-8) {
            viewDir.set(1, 0.72, 0.94);
        }
        viewDir.normalize();

        this.controls.target.copy(center);
        this.camera.position.copy(center).add(viewDir.multiplyScalar(distance));

        this.camera.near = Math.max(0.05, distance - (radius * 3));
        this.camera.far = Math.max(200, distance + (radius * 8));
        this.camera.updateProjectionMatrix();
        this.controls.update();
    }

    resetView() {
        this.fitCameraToObject(this.compositeGroup, 1.38, new THREE.Vector3(1, 0.72, 0.94));
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
        this.updateIntrinsicRightAngleMarkerVisibility();
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