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

function hasSharedStateInUrl() {
    const hash = window.location.hash || '';
    if (!hash.startsWith('#')) return false;
    const params = new URLSearchParams(hash.slice(1));
    return !!params.get(SHARE_HASH_KEY);
}

function restartTitleAnimation() {
    const titleGoodie = titleScreen.querySelector('.title-goodie');
    if (!titleGoodie) return;

    const resetGoodie = titleGoodie.cloneNode(true);
    titleGoodie.replaceWith(resetGoodie);
}

function returnToTitleScreen() {
    titleScreen.classList.remove('hidden');
    restartTitleAnimation();
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
window.addEventListener('pageshow', () => window.setTimeout(setActualVH, 0));

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
        if (trimensionApp?.isCrashReportOpen?.()) {
            trimensionApp.closeCrashReport();
            return;
        }

        if (trimensionApp?.isTriangleExtractionOpen?.()) {
            trimensionApp.closeTriangleExtraction();
            return;
        }

        if (helpOverlay.classList.contains('show')) {
            helpOverlay.classList.remove('show');
            return;
        }

        if (appInitialized) {
            returnToTitleScreen();
        }
    }
});

window.addEventListener('load', () => {
    if (!appInitialized && hasSharedStateInUrl()) {
        startApp();
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
        if (params.baseMirror) {
            return [
                new THREE.Vector3(base / 3, yPos, -triangleHeight / 3),
                new THREE.Vector3((-2 * base) / 3, yPos, -triangleHeight / 3),
                new THREE.Vector3(base / 3, yPos, (2 * triangleHeight) / 3)
            ];
        }

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

const SHARE_STATE_VERSION = 1;
const SHARE_HASH_KEY = 'state';
const MAX_SHARE_URL_LENGTH = 7000;

const BUILT_IN_EXAMPLES = [
    {
        name: 'AQA GCSE 2013',
        payload: 'g.H4sIAAAAAAAACr2W227jNhCGX8WYa9ogdZbunKTZBqibOFkEWRiGQUuszVYiBYrKJsj6ts_Up2hfqaCsYxKnbjfpTZAZkprvnwPpJ7hnquBSQEQQlByiJ8ipYOllzgREWpVshyCmGVO0WpMF19X2xTicYGz7IQk9BxPfwyFyJ9izHcsnIXYc18EeItbEcywHu2EYYt_xXWeJQFO1YRqiBZ4Qy_OdMAwDF4dBYIeITIIwIJbTugKEJ9i1sRf6BAcBdkmAneUOQUrXLC0M1JoWDKLFAgq8-gEQXMESLaAgq2ln4NU5IJg3KyedgVefAMF1s3LaGXj1IyC4aVbO-oaJc9ts-xkQ3DXGDBB8abZd7o0lgoQpfs-SijPjSS650N9y71vuA4IZLI2kWGZVgplRJdiDvkmlvkggshAUqdQFRIsn4AlEGEGueMY1v2cQQVyuJU8AgVScCU33JYJCU5FQZRZyqmhWZesrT_QWIt8g5eY_d-Ii2DK-2WqIbGNwEadlws5pzE6Z0EwVM5mYOFrm47XUWmawQ7CVRUsoyjTde8ypzkO1pvH2Nd-13HPOS6o0U59LJQqI8A7tFZKhQsViTcWmTKka54-KZq_IpTl7GJf5QG3KxKaWWwt_Kdecu2obG-JK8nOBeKjOpAL-izxTZ7n-lcW6aKp8WZnmA8Q3TSAKrcrYHD6VqVQXImEPEIUIuGZZ1wPEQ6Af86ouilOxSRkgEDQzns-1ZzQ9MR1dlGvNdVqtbBVj46r_RgWrwgCCe17wtVk3I29a4xcu6oQ8wW9cJMMo1fGLxMBATozDMn9sWBoBqVQGz7cd1w0QyJzGXD9ChCdWsGsL7Lb4z9inFfhfvw-5916qR7MjcV9nNbD1LI5fzGEeGgGaPWiIoAJoxLhGDrE7eueY5N_Ovnx48o8QVIuwsediA36gIHYrqWCbjAndKfpken9UlOtxvTQ6GQqb1ZH7e45U1u3uCzssyTeStjxJ2heqVWD9GwXTD1aQe4crc1ABaRVU71vH3-LNr4fcZ_sIoybCkczN9nFz6AX6HrPgG0F1qVjvTC2io8YtdZ5S0ZuEeUkTxVOqmaLp6Gp-fTOEP5elqschluYsVaOc6nh7pIom3ADeNY5aAYI86I2AF3qeQ8LhCLQ6wqNm-nkBPmimezXoblQS-N6h-Q3evk___OPAfTr97vt0_wIQ3Ls8q2g1d4hD1_G6NPtHpXl69_-kucfe5dmzHRcfSrR7aEZ_MuZoejYE90bZkaAs2bBx882XT6zTy6_5aAvkvA10Mh8C2e8AZO07tAOy-0D2P2To2eNB8HulyOoRVV9tkay3ke5uh0juexDhqrVIj8ntI5HDr9VN_UbdPh-Cr7K9Lrufid_3TJFXJ6D67bbbLXe7vwHYpt0qJA4AAA',
    },
    {
        name: 'AQA L2FM 2022',
        payload: 'g.H4sIAAAAAAAACrWW0W7iOBSGXwWdaxc5sZM4uaNI26k0o2Fneoe4MIkLnglOZDtdqpb3mqfYZ1rZQAKZUhht4So-ceLv__lP7Bd4EtrISkEWIGgkZC9QcyXKr7VQkFndiA2CnK-E5v5eZaT106fxME4IjeI0CGISxylF0ZDEJElCGsVhRFlMUYCHEUlDxmhMkyRkaTRDYLleCAvZFA-DkKTuxyKcMkZYjPAwoCwI6b5GMUF4iCOC4zQJMGM4Chim8WyDoORzURqHNedGQDadISiElk-igGw6Bc2trF7r5LVmr-Q1AAQTmLkH82rlhQj3rBJr-72s7H3hPTBlZQ1k0xeQBWQYQa3lSlr5JCCDvJlXsgAElZZCWb61AozlquDa3ai55ivP9I8s7BKyxDHV7ooiWAq5WFrIIgRS5WVTiL94LsZCWaHNl6pwa1SPj7BBsKxMi6WastxW3PSuwq3l-fKt2rdqC_d3w7UV-qHRykCGN059Nf8hcmv22r_6oVcfO2uUsbrJ3cPjqqz0vSrE2sNLK1adM0GEwD7Xjtj_D4BA8ZUbfnbDwWgMCEwzt9KWrvoMCJ6kkXM3crlyvjxKtYvTC_yUqoAMRLEQN_s31pVU9r5wq0Lt_sCagEuQWFv_ys0G7WhoS2O15GpRig7oYVcZjG57UA9LLcSNX2VghBd9IebBKr9Dhi1p7iyELMVpROMQQVXzXNpnyPAwZB09ed_LSQ97_X-93HXJze8dcmTw-sDg8BKDx3eTaxtMXCEBdFJD53oQJzGhET5pe3DG9tGxmDAa5KsrWh8cWL9dq0PFLWrP-5E3_t9fx6jbKreDyYW8J8N8YVQ8QOc7TQMStvTpRc3ZT_l1mvOcnk4DoVHETmWHnYnOXS867LrRSQ6jw46ik5wh_XRMGoTXJWUHpNu1WtK4JTVisRLKdqx3bvMbmGZ-s7s1uD3G_uZW2oelm3ahkG72n_i9lEXRnpL2IqI_ETG6voiavf-tfFMEPZUZTziYeMJPd2_QD4KMXMirO7VvMW899pNG_mTmL28hIwiMXChuGy32rznQ1Io4s6mO-y2KPyD4--3poBvxUcbD96H6B5T0I5jC_sc6Pd5bzhzkep1GP8Sn3TmpY6I7n2abzX9gc3HBlQwAAA',
    },
    {
        name: 'AQA GCSE 2018',
        payload: 'g.H4sIAAAAAAAACrWUzY7aMBSFXwXdtUG2Y4fEuwAz0kitRm1nh9DIJC64zZ8cg0CId-qy6z5An6lyICQwf0gz3SDOTbj387nH7GCtTKWLHARBsNIgdlDKXKX3pcpBWLNSewSxzJSR9bOi0rZ-fUrogPhkiMMh8UKP-BTxwZDhgIQsCEOP-yFB_eGAURpQz-ce5QEOyAyBlWahLIgpHfgkDIYsDAOOwyBgCA8ICwhtK9hDeIC5h_1wSHAQYE4CzGZ7BKmcq7RyTHNZKRDTKVTe4wgQRDBDtbgBBKNGjAHBTSMmgGDciAgQTA6C1U8eGtHpxuoG40ZEbWvefY13h_LuUN4dyh9v26G87nYLsxmCRBm9VgmIqTtjXGS14codM1cb-y0t7F0CwkdQpYWtQEx3oBMQHEFpdKatXisQYPRiafvWaJkvUtUvja4yQFAYrXIrDwuEyso8kSYBBKU0MqvNTNUiqtulajECQd2XfGGXIPiAI2hafi6Sdo6cF2vVj2CPYFlUJ8h8laaHyq2MVVuR1sp4-Vzta3GA-7KSxirzsDJ5BQLvnRfF_IeKbdU4cV9L18DznVF5Zc0qdj8eF2lh7vJEbUAQikBblbVGee4Q29LB1wcBBLnMnIyc7P39DQiq1dxqm7ZVaXtuR2td6bkru3vhlvVd58frsIOfOk86XctC5_YucZOh9F2Bug8G7gaojQUB9azY4YIIcciZT_d7dORkJ8466S3nJyd70ficc3slnkoWqt90PGN8greFlsY70TQJaIEejpVeNLqAelgapfr1lF6l6u1cidmZ8hSSn0iP5hF_6HuMYwRFKWNttyDwgAYtPr0Gf3Lp6cfjX-SgxWchcZAv4JPXM-v_-fVCaEfvDi2_TMVhWEvuMc47qPj12I4n56Sbd8eWuYLfAdy0saXhVXv__7H1n48td_4R76W90-ANM0fn3BzjXvYxhvKOoYe2LdXwjX-mCyr2IVT0koodqWb7_T-IAbQExQgAAA',
    },
    {
        name: 'AQA L2FM 2021',
        payload: 'g.H4sIAAAAAAAACs2X247bNhCGX0WYa65B6kjpzuttNwvUmzbxnWEUtMS12ciSQFEbb9e6zfPkGfICfaWCtGzJx3WbGIgvDHF0mO8f_kNRr_DMZSnyDCKCoBIQvULBMp6-L3gGkZIVrxHEbMElM-fyUihz-djvEc8Ng5BgEmLbpwQFPd8JPUy90PYIDkIbEbvnYuzggDjYIb5DgwkCxeSMK4jGuEdsJ9Q_6uGQUof6CPeIS4ntbmIudhDuYc_BfhgQTCn2CMWuP6kRpGzK01JjTVnJIRqPoSR_3gGCEUyQGfQBwWAzGACCPkwmCBIuxTNPzC2SKZGvCndV0JW9IoBgaG5o4s6qCJr4I0x02jhfmDJwnTnjS_UxzdVDApGDoExzVUI0fgWRQGQjKKRYCCWeOUQQV9NcJIAgl4Jniq0LCaViWcKkPlEwyRZG0WeRqDlEfs_TtMX2eM7FbK6agcjitEr4ryzmA54pLsthnuhM-dMT1AjmebmFy6o0XUf05W2EKcXi-bHYh3yN-EfFpOJyVMmshAjXugb59C8eq3JTgfdmqB9gU12grFSyivXNgzzN5UOW8CVEhCAQii86BQoQqJdCI5vJBAQZW-jhb3poPd4DgrKaKqFSHf0bEDyLUkz1SJtTF-dJZI0nX-GTyBKIgCczfrN5YpGLTD0kOutm4m8OJ7cIQHuTL5XJU9eoQfS3iEoKls1S3lKOmojVv73bJR3NJec3JrVVclOKC9k7WbrkhYG09Z-rSWNdWIiIH_iO62EEecFioV4gwj2btvje-Qrvc798b4U7iE0xXzrFdM_TDPdolj94vjtNvoO47CA6b1hyuIvo2Fa8uJ4tTwlo2dcALb-95d8za9849Z8vu_zrKFPW7YUijpnzTJkP7GAAWve6IXHslp5c0m3Dq3fbRYJaEY7refRUC5LwvKP6d1dwVHHg8T2fEPoG1e3VqOwzVG-8EPq_XI3KO0Pln-8pm377eqKrht_RVUVwZgE4OLNdNFoZDVfjUwf7HiZuq8q7pNfuh4_XfrP9b5mNMD_0fVd32Yn-a186JZ8teKZamR_XAWu45_bR57yR2N3MXKizzXLxkrKzmhAa-C29c5r-Xm_nrLKa3jSnrD0ZH3SuzVy1l_1IHd0tlKN1zEWSbL8dtirs_6Kif30Vx0zXtdZxFeTU6mQQrd8N4v3gCL5FIvtCYNnKPQa9rrK5qG--2szhrfnaKMUsY6qSfPOYjqhWBf65HXWwSTs6F-HPZih6frd2VMTJd3DXT-_uruYnuq3xxX7aaqrrSV3_C5fcWxJAEAAA',
    },
];

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
            vertexIds: ['A', 'B', 'C'],
            uAxis: (p) => {
                const [a, b, c] = getTriangularPrismProfilePoints(p, p.length / 2);
                const centroid = getTriangleCentroid(a, b, c);
                return c.clone().sub(centroid).normalize();
            },
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
            vertexIds: ['D', 'E', 'F'],
            uAxis: (p) => {
                const [a, b, c] = getTriangularPrismProfilePoints(p, -p.length / 2);
                const centroid = getTriangleCentroid(a, b, c);
                return c.clone().sub(centroid).normalize();
            },
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
            vertexIds: ['A', 'B', 'C'],
            uAxis: (p) => {
                const [baseA, baseB, baseC] = getTetrahedronBasePoints(p, -p.height / 2);
                const centroid = getTriangleCentroid(baseA, baseB, baseC);
                return baseC.clone().sub(centroid).normalize();
            },
            center: (p) => new THREE.Vector3(0, -p.height / 2, 0),
            dims: ['base', 'triangleHeight'],
            label: 'Base Triangle'
        },
    ]
};

class TrimensionApp {
    constructor() {
        this.canvas = document.getElementById('three-canvas');
        this.canvasContainer = this.canvas?.closest('.canvas-container') || this.canvas?.parentElement || null;
        this.canvasEmptyStateEl = null;
        if (this.canvasContainer) {
            const emptyState = document.createElement('div');
            emptyState.className = 'canvas-empty-state';
            emptyState.innerHTML = 'Use the <span class="inline-add-button" aria-hidden="true">+ Add</span> button to add up to 3 primitive objects to the diagram or choose an example diagram';
            this.canvasContainer.appendChild(emptyState);
            this.canvasEmptyStateEl = emptyState;
        }
        this.panelToggleBtn = document.getElementById('panel-toggle-btn');
        this.controlPanel = document.querySelector('.control-panel');
        this.pointsListEl = document.getElementById('points-list');
        this.selectionSummaryEl = document.getElementById('selection-summary');
        this.actionsListEl = document.getElementById('actions-list');
        this.objectSections = {
            triangles: { header: document.getElementById('triangles-section-header'), content: document.getElementById('triangles-section-content'), arrow: document.getElementById('triangles-section-arrow'), list: document.getElementById('triangles-list') },
            segments:  { header: document.getElementById('segments-section-header'),  content: document.getElementById('segments-section-content'),  arrow: document.getElementById('segments-section-arrow'),  list: document.getElementById('segments-list')  },
            angles:    { header: document.getElementById('angles-section-header'),    content: document.getElementById('angles-section-content'),    arrow: document.getElementById('angles-section-arrow'),    list: document.getElementById('angles-list')    },
            planes:    { header: document.getElementById('planes-section-header'),    content: document.getElementById('planes-section-content'),    arrow: document.getElementById('planes-section-arrow'),    list: document.getElementById('planes-list')    },
            labels:    { header: document.getElementById('labels-section-header'),    content: document.getElementById('labels-section-content'),    arrow: document.getElementById('labels-section-arrow'),    list: document.getElementById('labels-list')    },
        };
        Object.values(this.objectSections).forEach((section) => {
            section.title = section.header?.querySelector('h3') || null;
            section.baseTitle = section.title?.textContent || '';
        });
        this.primitiveSelect = document.getElementById('primitive-select');
        this.primitiveCardsListEl = document.getElementById('primitive-cards-list');
        this.primitiveChip = document.getElementById('primitive-chip');
        this.orientationChip = document.getElementById('orientation-chip');
        this.ghostToggleBtn = document.getElementById('ghost-toggle-btn');
        this.ghostWireIcon = document.getElementById('ghost-wire-icon');
        this.ghostSolidIcon = document.getElementById('ghost-solid-icon');
        this.labelBadgeToggleBtn = document.getElementById('label-badge-toggle-btn');
        this.labelPlainIcon = document.getElementById('label-plain-icon');
        this.labelBadgeIcon = document.getElementById('label-badge-icon');
        this.labelOffIcon = document.getElementById('label-off-icon');
        this.pointMarkerToggleBtn = document.getElementById('point-marker-toggle-btn');
        this.pointFilledIcon = document.getElementById('point-filled-icon');
        this.pointHollowIcon = document.getElementById('point-hollow-icon');
        this.gridToggleBtn = document.getElementById('grid-toggle-btn');
        this.gridIcon = document.getElementById('grid-icon');
        this.displaySizeToggleBtn = document.getElementById('display-size-toggle-btn');
        this.sizeSmallOption = document.getElementById('size-small');
        this.sizeLargeOption = document.getElementById('size-large');
        this.themeToggleBtn = document.getElementById('theme-toggle');
        this.lightIcon = document.getElementById('light-icon');
        this.darkIcon = document.getElementById('dark-icon');
        this.shareBtn = document.getElementById('share-button');
        this.addBtn = document.getElementById('add-btn');
        this.addDropdown = document.getElementById('add-dropdown');
        this.primitiveSectionHeader = document.getElementById('primitive-section-header');
        this.primitiveSectionContent = document.getElementById('primitive-section-content');
        this.primitiveSectionArrow = document.getElementById('primitive-section-arrow');
        this.primitiveSectionTitle = this.primitiveSectionHeader?.querySelector('h3') || null;
        this.primitiveSectionBaseTitle = this.primitiveSectionTitle?.textContent || '';
        this.pointsSectionHeader = document.getElementById('points-section-header');
        this.pointsSectionContent = document.getElementById('points-section-content');
        this.pointsSectionArrow = document.getElementById('points-section-arrow');
        this.pointsSectionTitle = this.pointsSectionHeader?.querySelector('h3') || null;
        this.pointsSectionBaseTitle = this.pointsSectionTitle?.textContent || '';
        this.triangleExtractOverlay = document.getElementById('triangle-extract-overlay');
        this.triangleExtractModal = document.getElementById('triangle-extract-modal');
        this.triangleExtractTitle = document.getElementById('triangle-extract-title');
        this.triangleExtractSubtitle = document.getElementById('triangle-extract-subtitle');
        this.triangleExtractFlightSvg = document.getElementById('triangle-extract-flight-svg');
        this.triangleExtractFlightPolygon = document.getElementById('triangle-extract-flight-polygon');
        this.triangleExtractFlightOutline = document.getElementById('triangle-extract-flight-outline');
        this.triangleExtractSvg = document.getElementById('triangle-extract-svg');
        this.triangleExtractStage = document.getElementById('triangle-extract-stage');
        this.triangleExtractPolygon = document.getElementById('triangle-extract-polygon');
        this.triangleExtractOutline = document.getElementById('triangle-extract-outline');
        this.triangleExtractRightAngle = document.getElementById('triangle-extract-right-angle');
        this.triangleExtractRotateBtn = document.getElementById('triangle-extract-rotate');
        this.triangleExtractFlipBtn = document.getElementById('triangle-extract-flip');
        this.triangleExtractCloseBtn = document.getElementById('triangle-extract-close');
        this.triangleExtractLabelEls = {
            A: document.getElementById('triangle-extract-label-a'),
            B: document.getElementById('triangle-extract-label-b'),
            C: document.getElementById('triangle-extract-label-c'),
            D: document.getElementById('triangle-extract-label-d')
        };
        this.triangleExtractSideEls = {
            AB: document.getElementById('triangle-extract-side-ab'),
            BC: document.getElementById('triangle-extract-side-bc'),
            CA: document.getElementById('triangle-extract-side-ca'),
            CD: document.getElementById('triangle-extract-side-cd'),
            DA: document.getElementById('triangle-extract-side-da')
        };
        this.triangleExtractAnglesGroup = document.getElementById('triangle-extract-angles-group');

        this.panelOpen = true;
        this.ghostFaces = true;
        this.pointMarkersVisible = true;
        this.labelMode = 'badge';
        this.gridVisible = true;
        this.displaySizeMode = this.getInitialDisplaySizeMode();
        this.themeMode = 'light';
        this.examplesAccordionCollapsed = true;
        this.nextObjectId = 1;
        this.sceneObjects = [];
        this.selectedPoints = [];
        this.pointDefinitions = [];
        this.derivedPoints = [];
        this.pointsHintDismissed = false;
        this.isRestoringSharedState = false;
        this.baseLabelOverrides = new Map();
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
        this.pointsSectionCollapsed = false;
        this.activeTriangleExtraction = null;
        this.triangleExtractTransitionState = 'closed';
        this.lastFocusedElementBeforeTriangleExtract = null;
        this.triangleExtractSettleTimer = null;
        this.triangleExtractAnimationFrame = null;
        this.crashReportOverlay = document.getElementById('crash-report-overlay');
        this.crashReportPre = document.getElementById('crash-report-content');
        this.crashReportRefreshBtn = document.getElementById('crash-report-refresh');
        this.crashReportCopyBtn = document.getElementById('crash-report-copy');
        this.crashReportCloseBtn = document.getElementById('crash-report-close');
        this.crashReportEntries = [];
        this.maxCrashReportEntries = 140;
        this.crashReportOpenedAt = null;
        this.crashReportShortcut = 'Ctrl+Alt+Shift+K';
        this._crashListenersBound = false;
        this.crashWatchdogIntervalMs = 30000;
        this.crashWatchdogIntervalId = null;
        this.wasDiscardedAtLoad = document.wasDiscarded === true;
        this.lastIntegrityDigest = null;

        this.defaultParams = {
            cuboid: { width: 7, depth: 4, height: 5, includeFaceCentersMode: 'off' },
            'right-triangle-prism': { legA: 5, legB: 4, length: 7, triangleMode: 'isosceles' },
            'rectangular-pyramid': { length: 7, width: 5, height: 6, apexPosition: 'center' },
            tetrahedron: { base: 6, triangleHeight: 4.5, height: 6, baseTriangleMode: 'isosceles', apexPosition: 'A', baseMirror: false },
            sphere: { radius: 3 },
            hemisphere: { radius: 3 },
            cylinder: { radius: 2.5, height: 6 },
            cone: { radius: 2.5, height: 6 }
        };

        // compositeSlots: array of { id, primitive, orientation, params, hostSlotId, hostFaceId, attachFaceId, attachRotationQuarterTurns }
        this.compositeSlots = [];
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
                    { key: 'legA', label: 'Width', min: 2, max: 10, step: 0.5 },
                    { key: 'legB', label: 'Height', min: 2, max: 10, step: 0.5 },
                    { key: 'length', label: 'Length', min: 2, max: 12, step: 0.5 }
                ]
            },
            'rectangular-pyramid': {
                label: 'Rectangular Pyramid',
                params: [
                    { key: 'length', label: 'Length', min: 2, max: 12, step: 0.5 },
                    { key: 'width', label: 'Width', min: 2, max: 12, step: 0.5 },
                    { key: 'height', label: 'Height', min: 2, max: 10, step: 0.5 }
                ]
            },
            tetrahedron: {
                label: 'Tetrahedron',
                params: [
                    { key: 'base', label: 'BASE', min: 2, max: 10, step: 0.5 },
                    { key: 'triangleHeight', label: 'HEIGHT', min: 2, max: 10, step: 0.5 },
                    { key: 'height', label: 'APEX', min: 2, max: 10, step: 0.5 }
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
                { key: 'plane', label: 'Add Quadrilateral' }
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

        this.setupCrashDiagnostics();
        this.initThree();
        this.bindEvents();
        this.buildComposite();
        this.renderCompositeCards();
        this.restoreStateFromUrlIfPresent();
        this.animate();
    }

    async restoreStateFromUrlIfPresent() {
        const payload = this.getSharePayloadFromHash();
        if (!payload) return;

        try {
            const snapshot = await this.decodeShareState(payload);
            if (!snapshot || snapshot.version !== SHARE_STATE_VERSION || !this.applySharedStateSnapshot(snapshot)) {
                throw new Error('Unsupported or invalid shared state');
            }
            this.applySharedUrlDefaultSectionState();
            this.showToast('Loaded shared diagram.');
        } catch (error) {
            console.error('Failed to restore shared state:', error);
            this.showAlertModal('Unable to load the shared diagram state from this URL.');
        }
    }

    async loadBuiltInExample(index) {
        const example = BUILT_IN_EXAMPLES[index];
        if (!example) return;
        try {
            const snapshot = await this.decodeShareState(example.payload);
            if (!snapshot || snapshot.version !== SHARE_STATE_VERSION || !this.applySharedStateSnapshot(snapshot)) {
                throw new Error('Invalid example snapshot');
            }
            this.applyBuiltInExampleSectionState();
            this.showToast(`Loaded: ${example.name}`);
        } catch (error) {
            console.error('Failed to load built-in example:', error);
            this.showAlertModal(`Unable to load example: ${example.name}`);
        }
    }

    initThree() {
        this.scene = new THREE.Scene();
        this.scene.background = new THREE.Color(0xffffff);
        document.documentElement.setAttribute('data-theme', this.themeMode);
        this.scene.background.set(this.themeMode === 'dark' ? 0x606060 : 0xffffff);

        this.camera = new THREE.PerspectiveCamera(55, this.canvas.clientWidth / this.canvas.clientHeight, 0.01, 500);
        this.camera.position.set(10, 8, 11);

        const userAgent = navigator.userAgent || '';
        const isIPhoneOrIPod = /iPhone|iPod/.test(userAgent);
        const isAndroid = /Android/.test(userAgent);
        const isiPadOSDesktopUA = navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
        const isIPad = /iPad/.test(userAgent) || isiPadOSDesktopUA;
        const isAndroidTablet = isAndroid && !/Mobile/.test(userAgent);
        const isTablet = isIPad || /tablet/i.test(userAgent) || isAndroidTablet;
        const isMobilePhone = (isIPhoneOrIPod || (isAndroid && /Mobile/.test(userAgent))) && !isTablet;

        this.renderer = new THREE.WebGLRenderer({ canvas: this.canvas, antialias: true, alpha: false, logarithmicDepthBuffer: true });
        const pixelRatio = isMobilePhone
            ? Math.min(window.devicePixelRatio || 1, 2)
            : (window.devicePixelRatio || 1);
        this.renderer.setPixelRatio(pixelRatio);
        this.renderer.setSize(this.canvas.clientWidth, this.canvas.clientHeight, false);

        this.controls = new OrbitControls(this.camera, this.canvas);
        this.controls.enableDamping = true;
        this.controls.dampingFactor = 0.08;
        this.controls.target.set(0, 0, 0);
        this.controls.minDistance = isMobilePhone ? 2 : 1;
        this.controls.maxDistance = 100;
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
        this.grid.visible = this.gridVisible;
        this.scene.add(this.grid);
        this.updateGridThemeAppearance();

        window.addEventListener('resize', this.handleWindowResize);
    }

    bindEvents() {
        this.panelToggleBtn.addEventListener('click', () => {
            this.panelOpen = !this.panelOpen;
            this.controlPanel.classList.toggle('closed', !this.panelOpen);
            this.panelToggleBtn.classList.toggle('active', this.panelOpen);
        });

        this.canvas.addEventListener('pointerdown', this.handleCanvasPointerDown, { passive: true });

        this._keysHeld = new Set();
        this._handleKeyDown = (e) => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) return;
            const isCrashReportShortcut = e.ctrlKey && e.altKey && e.shiftKey && e.code === 'KeyK';
            if (isCrashReportShortcut) {
                e.preventDefault();
                this.toggleCrashReport();
                return;
            }

            if (this.isCrashReportOpen()) {
                if (e.key === 'Escape') {
                    e.preventDefault();
                    this.closeCrashReport();
                }
                return;
            }

            if (this.triangleExtractOverlay?.classList.contains('show')) return;
            const nav = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'];
            if (nav.includes(e.key)) e.preventDefault();
            this._keysHeld.add(e.code);
        };
        this._handleKeyUp = (e) => { this._keysHeld.delete(e.code); };
        window.addEventListener('keydown', this._handleKeyDown);
        window.addEventListener('keyup', this._handleKeyUp);

        if (this.triangleExtractCloseBtn) {
            this.triangleExtractCloseBtn.addEventListener('click', () => {
                if (this.triangleExtractTransitionState === 'open') {
                    this.closeTriangleExtraction();
                }
            });
        }

        if (this.triangleExtractRotateBtn) {
            this.triangleExtractRotateBtn.addEventListener('click', () => this.rotateTriangleExtractionLayout());
        }

        if (this.triangleExtractFlipBtn) {
            this.triangleExtractFlipBtn.addEventListener('click', () => this.flipTriangleExtractionLayout());
        }

        if (this.triangleExtractOverlay) {
            this.triangleExtractOverlay.addEventListener('click', (event) => {
                if (event.target === this.triangleExtractOverlay) return;
            });
        }

        if (this.crashReportCopyBtn) {
            this._handleCrashReportCopyClick = () => this.copyCrashReport();
            this.crashReportCopyBtn.addEventListener('click', this._handleCrashReportCopyClick);
        }

        if (this.crashReportRefreshBtn) {
            this._handleCrashReportRefreshClick = () => this.refreshCrashReport();
            this.crashReportRefreshBtn.addEventListener('click', this._handleCrashReportRefreshClick);
        }

        if (this.crashReportCloseBtn) {
            this._handleCrashReportCloseClick = () => this.closeCrashReport();
            this.crashReportCloseBtn.addEventListener('click', this._handleCrashReportCloseClick);
        }

        if (this.crashReportOverlay) {
            this._handleCrashReportOverlayClick = (event) => {
                if (event.target === this.crashReportOverlay) {
                    this.closeCrashReport();
                }
            };
            this.crashReportOverlay.addEventListener('click', this._handleCrashReportOverlayClick);
        }

        this.addBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            this.examplesAccordionCollapsed = true;
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
                    item.innerHTML = `<strong>+</strong> ${this.primitiveMeta[primKey].label}`;
                    this.addDropdown.appendChild(item);
                });
            }

            const divider = document.createElement('div');
            divider.className = 'dropdown-divider';
            this.addDropdown.appendChild(divider);

            const accordionBtn = document.createElement('button');
            accordionBtn.type = 'button';
            accordionBtn.className = 'dropdown-accordion-header';
            accordionBtn.setAttribute('data-examples-accordion', '');
            accordionBtn.setAttribute('aria-expanded', String(!this.examplesAccordionCollapsed));
            accordionBtn.innerHTML = '<span class="dropdown-accordion-arrow">&#9654;&#xFE0E;</span><span class="dropdown-accordion-label">Examples</span>';
            this.addDropdown.appendChild(accordionBtn);

            const accordionContent = document.createElement('div');
            accordionContent.className = 'dropdown-accordion-content' + (this.examplesAccordionCollapsed ? ' collapsed' : '');
            accordionContent.setAttribute('data-examples-content', '');
            BUILT_IN_EXAMPLES.forEach((example, idx) => {
                const item = document.createElement('div');
                item.className = 'dropdown-item dropdown-item-example';
                item.dataset.example = String(idx);
                item.textContent = example.name;
                accordionContent.appendChild(item);
            });
            this.addDropdown.appendChild(accordionContent);

            const isOpening = this.addDropdown.style.display === 'none';
            this.addDropdown.style.display = isOpening ? 'block' : 'none';
        });

        this.addDropdown.addEventListener('click', (event) => {
            event.stopPropagation();
            const accordionHeader = event.target.closest('[data-examples-accordion]');
            if (accordionHeader) {
                this.examplesAccordionCollapsed = !this.examplesAccordionCollapsed;
                accordionHeader.setAttribute('aria-expanded', String(!this.examplesAccordionCollapsed));
                const content = this.addDropdown.querySelector('[data-examples-content]');
                if (content) { content.classList.toggle('collapsed', this.examplesAccordionCollapsed); }
                return;
            }
            const primitiveItem = event.target.closest('[data-primitive]');
            if (primitiveItem) {
                this.addSlot(primitiveItem.dataset.primitive);
                this.expandPrimarySections();
                this.addDropdown.style.display = 'none';
                return;
            }
            const exampleItem = event.target.closest('[data-example]');
            if (exampleItem) {
                this.addDropdown.style.display = 'none';
                this.loadBuiltInExample(Number(exampleItem.dataset.example));
            }
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

            const prismAttachFaceBtn = event.target.closest('[data-cycle-prism-attach-face-slot-id]');
            if (prismAttachFaceBtn) {
                this.cyclePrismAttachFace(Number(prismAttachFaceBtn.dataset.cyclePrismAttachFaceSlotId));
                return;
            }

            const prismAttachRotationBtn = event.target.closest('[data-cycle-prism-attach-rotation-slot-id]');
            if (prismAttachRotationBtn) {
                this.cyclePrismAttachRotation(Number(prismAttachRotationBtn.dataset.cyclePrismAttachRotationSlotId));
                return;
            }

            const triangleModeBtn = event.target.closest('[data-cycle-triangle-mode-slot-id]');
            if (triangleModeBtn) {
                this.cycleTriangularPrismMode(Number(triangleModeBtn.dataset.cycleTriangleModeSlotId));
                return;
            }

            const cuboidFaceCentersBtn = event.target.closest('[data-toggle-cuboid-face-centers-slot-id]');
            if (cuboidFaceCentersBtn) {
                this.toggleCuboidFaceCenters(Number(cuboidFaceCentersBtn.dataset.toggleCuboidFaceCentersSlotId));
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
                const extractButton = event.target.closest('[data-extract-object-id]');
                const editLabelButton = event.target.closest('[data-edit-label-object-id]');
                const editAngleButton = event.target.closest('[data-edit-angle-object-id]');
                const toggleButton = event.target.closest('[data-toggle-object-id]');
                const deleteButton = event.target.closest('[data-delete-object-id]');
                if (extractButton) {
                    this.openTriangleExtraction(Number(extractButton.dataset.extractObjectId));
                    this.closePanelOnMobile();
                    return;
                }
                if (editLabelButton) {
                    this.editLabelFromObjectCard(Number(editLabelButton.dataset.editLabelObjectId));
                    return;
                }
                if (editAngleButton) {
                    this.editAngleFromObjectCard(Number(editAngleButton.dataset.editAngleObjectId));
                    return;
                }
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

        const clearObjectsBtn = document.getElementById('clear-objects-btn');
        if (clearObjectsBtn) {
            clearObjectsBtn.addEventListener('click', () => this.clearAllObjects());
        }
        document.getElementById('reset-view-btn').addEventListener('click', () => {
            this.resetView();
            this.closePanelOnMobile();
        });
        this.ghostToggleBtn.addEventListener('click', () => {
            this.ghostFaces = !this.ghostFaces;
            this.updatePrimitiveMaterial();
            this.updateGhostToggleUI();
        });
        this.updateGhostToggleUI();

        if (this.labelBadgeToggleBtn) {
            this.labelBadgeToggleBtn.addEventListener('click', () => this.toggleLabelBadgeMode());
            this.updateLabelBadgeToggleUI();
        }

        if (this.pointMarkerToggleBtn) {
            this.pointMarkerToggleBtn.addEventListener('click', () => this.togglePointMarkers());
            this.updatePointMarkerToggleUI();
        }

        if (this.displaySizeToggleBtn) {
            this.displaySizeToggleBtn.addEventListener('click', () => this.toggleDisplaySizeMode());
            this.updateDisplaySizeToggleUI();
        }

        if (this.themeToggleBtn) {
            this.themeToggleBtn.addEventListener('click', () => this.toggleThemeMode());
            this.updateThemeToggleUI();
        }

        if (this.gridToggleBtn) {
            this.gridToggleBtn.addEventListener('click', () => this.toggleGrid());
            this.updateGridToggleUI();
        }

        if (this.shareBtn) {
            this.shareBtn.addEventListener('click', () => {
                this.handleShareButtonClick();
            });
        }

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

    expandPrimarySections() {
        this.primitiveSectionCollapsed = false;
        this.primitiveSectionContent.classList.remove('collapsed');
        this.primitiveSectionHeader.setAttribute('aria-expanded', 'true');
        this.primitiveSectionArrow.textContent = '\u25BC\uFE0E';

        this.pointsSectionCollapsed = false;
        this.pointsSectionContent.classList.remove('collapsed');
        this.pointsSectionHeader.setAttribute('aria-expanded', 'true');
        this.pointsSectionArrow.textContent = '\u25BC\uFE0E';
    }

    handleCanvasPointerDown(event) {
        if (event.target.closest('.control-panel')) {
            return;
        }

        if (this.shouldAutoClosePanelOnCanvasTap() && this.panelOpen) {
            this.closePanelOnMobile();
        }
    }

    shouldAutoClosePanelOnCanvasTap() {
        const isPhoneNarrow = window.innerWidth < 768;
        const userAgent = navigator.userAgent || '';
        const isiPadOSDesktopUA = navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
        const isIPad = /iPad/.test(userAgent) || isiPadOSDesktopUA;
        const isIPadPortrait = isIPad && window.innerHeight > window.innerWidth;

        return isPhoneNarrow || isIPadPortrait;
    }

    getInitialDisplaySizeMode() {
        const userAgent = navigator.userAgent || '';
        const isMobileUA = /iPhone|iPod|Android.*Mobile|Windows Phone|webOS|BlackBerry|IEMobile|Opera Mini/i.test(userAgent);
        const isNarrowViewport = window.innerWidth < 768;

        return (isNarrowViewport || isMobileUA) ? 'large' : 'small';
    }

    closePanelOnMobile() {
        if (!this.shouldAutoClosePanelOnCanvasTap() || !this.panelOpen) {
            return;
        }

        this.panelOpen = false;
        this.controlPanel.classList.add('closed');
        this.panelToggleBtn.classList.remove('active');
    }

    base64UrlEncode(bytes) {
        let binary = '';
        for (let i = 0; i < bytes.length; i += 1) {
            binary += String.fromCharCode(bytes[i]);
        }

        return btoa(binary)
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=+$/g, '');
    }

    base64UrlDecode(input) {
        const normalized = input
            .replace(/-/g, '+')
            .replace(/_/g, '/');
        const padding = normalized.length % 4 === 0 ? '' : '='.repeat(4 - (normalized.length % 4));
        const binary = atob(normalized + padding);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i += 1) {
            bytes[i] = binary.charCodeAt(i);
        }
        return bytes;
    }

    async streamToUint8Array(stream) {
        const response = new Response(stream);
        const buffer = await response.arrayBuffer();
        return new Uint8Array(buffer);
    }

    async encodeShareState(snapshot) {
        const json = JSON.stringify(snapshot);
        const rawBytes = new TextEncoder().encode(json);

        if (typeof CompressionStream === 'function') {
            try {
                const compressed = await this.streamToUint8Array(
                    new Blob([rawBytes]).stream().pipeThrough(new CompressionStream('gzip'))
                );
                if (compressed.length < rawBytes.length) {
                    return `g.${this.base64UrlEncode(compressed)}`;
                }
            } catch {
                // Fallback to raw payload.
            }
        }

        return `r.${this.base64UrlEncode(rawBytes)}`;
    }

    async decodeShareState(payload) {
        if (typeof payload !== 'string' || !payload.includes('.')) {
            throw new Error('Invalid payload format');
        }

        const [encoding, data] = payload.split('.', 2);
        const bytes = this.base64UrlDecode(data);

        if (encoding === 'g') {
            if (typeof DecompressionStream !== 'function') {
                throw new Error('Compressed payload unsupported in this browser');
            }

            const decompressed = await this.streamToUint8Array(
                new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'))
            );
            return JSON.parse(new TextDecoder().decode(decompressed));
        }

        if (encoding === 'r') {
            return JSON.parse(new TextDecoder().decode(bytes));
        }

        throw new Error('Unknown payload encoding');
    }

    getSharePayloadFromHash() {
        const hash = window.location.hash || '';
        if (!hash.startsWith('#')) return null;
        const hashBody = hash.slice(1);
        const params = new URLSearchParams(hashBody);
        return params.get(SHARE_HASH_KEY);
    }

    applySectionCollapseStateToUI() {
        this.primitiveSectionContent.classList.toggle('collapsed', this.primitiveSectionCollapsed);
        this.primitiveSectionHeader.setAttribute('aria-expanded', this.primitiveSectionCollapsed ? 'false' : 'true');
        this.primitiveSectionArrow.textContent = this.primitiveSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';

        this.pointsSectionContent.classList.toggle('collapsed', this.pointsSectionCollapsed);
        this.pointsSectionHeader.setAttribute('aria-expanded', this.pointsSectionCollapsed ? 'false' : 'true');
        this.pointsSectionArrow.textContent = this.pointsSectionCollapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';

        Object.entries(this.objectSections).forEach(([key, sec]) => {
            const collapsed = !!this.objectGroupCollapsed[key];
            sec.content.classList.toggle('collapsed', collapsed);
            sec.header.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
            sec.arrow.textContent = collapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
        });
    }

    normalizeSlotForRestore(slot) {
        if (!slot || typeof slot !== 'object') return null;
        if (typeof slot.primitive !== 'string' || !this.defaultParams[slot.primitive]) return null;

        const orientationOptions = this.orientations[slot.primitive] || [{ value: 'standard' }];
        const fallbackOrientation = orientationOptions[0]?.value || 'standard';
        const orientation = orientationOptions.some((opt) => opt.value === slot.orientation)
            ? slot.orientation
            : fallbackOrientation;

        const defaultParams = this.defaultParams[slot.primitive] || {};
        const params = { ...defaultParams, ...(slot.params || {}) };

        return {
            id: Number.isFinite(slot.id) ? Number(slot.id) : null,
            primitive: slot.primitive,
            orientation,
            params,
            hostSlotId: Number.isFinite(slot.hostSlotId) ? Number(slot.hostSlotId) : null,
            hostFaceId: typeof slot.hostFaceId === 'string' ? slot.hostFaceId : null,
            attachFaceId: typeof slot.attachFaceId === 'string' ? slot.attachFaceId : null,
            attachRotationQuarterTurns: Number.isFinite(slot.attachRotationQuarterTurns) ? Number(slot.attachRotationQuarterTurns) : 0
        };
    }

    applySharedUrlDefaultSectionState() {
        this.panelOpen = true;
        this.primitiveSectionCollapsed = false;
        this.pointsSectionCollapsed = false;
        this.objectGroupCollapsed = {
            triangles: true,
            segments: true,
            angles: true,
            planes: true,
            labels: true
        };

        this.panelToggleBtn.classList.add('active');
        this.controlPanel.classList.remove('closed');
        this.applySectionCollapseStateToUI();
    }

    applyBuiltInExampleSectionState() {
        this.panelOpen = true;
        this.primitiveSectionCollapsed = false;
        this.pointsSectionCollapsed = false;
        this.objectGroupCollapsed = {
            triangles: false,
            segments: true,
            angles: true,
            planes: true,
            labels: true
        };

        this.panelToggleBtn.classList.add('active');
        this.controlPanel.classList.remove('closed');
        this.applySectionCollapseStateToUI();
    }

    getShareableStateSnapshot() {
        return {
            version: SHARE_STATE_VERSION,
            ui: {
                panelOpen: this.panelOpen
            },
            camera: {
                position: [this.camera.position.x, this.camera.position.y, this.camera.position.z],
                target: [this.controls.target.x, this.controls.target.y, this.controls.target.z]
            },
            labels: {
                base: Array.from(this.baseLabelOverrides.entries()),
                derived: Array.from(this.derivedLabelOverrides.entries())
            },
            composite: {
                nextSlotId: this.nextSlotId,
                slots: this.compositeSlots.map((slot) => ({
                    id: slot.id,
                    primitive: slot.primitive,
                    orientation: slot.orientation,
                    params: { ...slot.params },
                    hostSlotId: slot.hostSlotId,
                    hostFaceId: slot.hostFaceId,
                    attachFaceId: slot.attachFaceId ?? null,
                    attachRotationQuarterTurns: slot.attachRotationQuarterTurns ?? 0
                }))
            },
            objects: {
                nextObjectId: this.nextObjectId,
                constructionColorIndex: this.constructionColorIndex,
                items: this.sceneObjects.map((item) => ({
                    id: item.id,
                    type: item.type,
                    name: item.name,
                    subtitle: item.subtitle,
                    visible: item.visible,
                    definition: item.definition ? JSON.parse(JSON.stringify(item.definition)) : null
                }))
            }
        };
    }

    applySharedStateSnapshot(snapshot) {
        if (!snapshot || typeof snapshot !== 'object') {
            return false;
        }

        const ui = snapshot.ui || {};
        this.panelOpen = ui.panelOpen !== false;

        this.panelToggleBtn.classList.toggle('active', this.panelOpen);
        this.controlPanel.classList.toggle('closed', !this.panelOpen);

        this.baseLabelOverrides = new Map(Array.isArray(snapshot.labels?.base) ? snapshot.labels.base : []);
        this.derivedLabelOverrides = new Map(Array.isArray(snapshot.labels?.derived) ? snapshot.labels.derived : []);

        this.isRestoringSharedState = true;
        try {

        const rawSlots = Array.isArray(snapshot.composite?.slots) ? snapshot.composite.slots : [];
        const normalizedSlots = rawSlots
            .map((slot) => this.normalizeSlotForRestore(slot))
            .filter(Boolean)
            .slice(0, 3);

        if (normalizedSlots.length > 0) {
            const usedIds = new Set();
            let nextFallbackId = 0;
            normalizedSlots.forEach((slot) => {
                if (slot.id == null || usedIds.has(slot.id)) {
                    while (usedIds.has(nextFallbackId)) nextFallbackId += 1;
                    slot.id = nextFallbackId;
                }
                usedIds.add(slot.id);
                nextFallbackId = Math.max(nextFallbackId, slot.id + 1);
            });
            this.compositeSlots = normalizedSlots;
            const candidateNextSlotId = Number(snapshot.composite?.nextSlotId);
            this.nextSlotId = Number.isFinite(candidateNextSlotId)
                ? Math.max(candidateNextSlotId, ...normalizedSlots.map((slot) => slot.id + 1))
                : Math.max(...normalizedSlots.map((slot) => slot.id + 1));
        }

        this.resetSceneObjects();
        this.buildComposite({ fitCamera: false });
        this.renderCompositeCards();

        // Two-pass restore: non-labels first so derived points (midpoints, ratio points) exist
        // before labels try to resolve their point IDs and attachment checks.
        const NON_LABEL_ORDER = { segment: 0, triangle: 1, angle: 2, plane: 3 };
        const allSavedObjects = Array.isArray(snapshot.objects?.items) ? snapshot.objects.items : [];
        const isDeferredVisualLabel = (saved) => {
            const kind = saved?.definition?.kind;
            return kind === 'edge-label' || kind === 'length-label' || kind === 'point-label';
        };
        const nonLabelObjects = allSavedObjects
            .filter((s) => !isDeferredVisualLabel(s))
            .sort((a, b) => (NON_LABEL_ORDER[a?.type] ?? 3) - (NON_LABEL_ORDER[b?.type] ?? 3));
        const derivedPointHelperObjects = nonLabelObjects.filter((s) => s?.definition?.kind === 'midpoint-point' || s?.definition?.kind === 'ratio-point');
        const otherNonLabelObjects = nonLabelObjects.filter((s) => s?.definition?.kind !== 'midpoint-point' && s?.definition?.kind !== 'ratio-point');
        const labelObjects = allSavedObjects.filter((s) => isDeferredVisualLabel(s));

        this.sceneObjects = [];
        let maxObjectId = 0;

        const restoreOne = (saved) => {
            if (!saved || !saved.definition) return false;
            const definition = JSON.parse(JSON.stringify(saved.definition));
            const object3D = this.createObjectFromDefinition(definition);
            if (!object3D) return false;

            const id = Number.isFinite(saved.id) ? Number(saved.id) : (maxObjectId + 1);
            const visible = saved.visible !== false;
            const type = typeof saved.type === 'string' ? saved.type : 'segment';
            const name = typeof saved.name === 'string' ? saved.name : this.getSceneObjectDisplayName({ definition, name: type });
            const subtitle = typeof saved.subtitle === 'string' ? saved.subtitle : '';

            object3D.userData.sceneObjectId = id;
            object3D.visible = visible;
            this.scene.add(object3D);
            this.sceneObjects.push({ id, type, name, subtitle, object3D, definition, visible });
            maxObjectId = Math.max(maxObjectId, id);
            return true;
        };

        const restorePendingObjects = (items, options = {}) => {
            let pending = items.slice();
            const maxPasses = 8;

            for (let pass = 0; pass < maxPasses && pending.length > 0; pass += 1) {
                let restoredThisPass = 0;
                const remaining = [];

                pending.forEach((saved) => {
                    if (restoreOne(saved)) {
                        restoredThisPass += 1;
                        return;
                    }
                    remaining.push(saved);
                });

                pending = remaining;
                if (restoredThisPass === 0) {
                    break;
                }

                if (options.refreshDerived !== false) {
                    this.refreshDerivedPoints();
                }
            }

            return pending;
        };

        // Pass 1: derived point helpers first so derived point IDs can materialise.
        restorePendingObjects(derivedPointHelperObjects);

        // Pass 2: geometry and other non-label objects.
        restorePendingObjects(otherNonLabelObjects);

        // Materialise derived points again after all geometry is restored.
        this.refreshDerivedPoints();

        // Pass 3: edge labels, point labels (now have backing geometry + derived points)
        restorePendingObjects(labelObjects, { refreshDerived: false });

        const nextObjectId = Number(snapshot.objects?.nextObjectId);
        this.nextObjectId = Number.isFinite(nextObjectId) ? Math.max(nextObjectId, maxObjectId + 1) : (maxObjectId + 1);
        const constructionColorIndex = Number(snapshot.objects?.constructionColorIndex);
        if (Number.isFinite(constructionColorIndex)) {
            this.constructionColorIndex = Math.max(0, constructionColorIndex);
        }

        this.pruneOrphanedSceneObjects();
        this.refreshDerivedPoints();
        this.buildPointMarkers();
        this.renderObjectsList();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
        this.updatePrimitiveMaterial();
        this.updateGhostToggleUI();
        this.updatePointMarkerToggleUI();
        this.updateLabelBadgeToggleUI();
        this.updateDisplaySizeToggleUI();
        this.applyThemeMode();
        if (this.grid) {
            this.grid.visible = this.gridVisible;
        }
        this.updateGridToggleUI();

        const cameraPos = Array.isArray(snapshot.camera?.position) ? snapshot.camera.position : null;
        const cameraTarget = Array.isArray(snapshot.camera?.target) ? snapshot.camera.target : null;
        if (cameraPos?.length === 3 && cameraTarget?.length === 3
            && cameraPos.every(Number.isFinite) && cameraTarget.every(Number.isFinite)) {
            this.camera.position.set(cameraPos[0], cameraPos[1], cameraPos[2]);
            this.controls.target.set(cameraTarget[0], cameraTarget[1], cameraTarget[2]);
            this.controls.update();
        } else {
            this.fitCameraToObject(this.compositeGroup, 1.38, new THREE.Vector3(1, 0.72, 0.94));
        }

        return true;
        } finally {
            this.isRestoringSharedState = false;
        }
    }

    async handleShareButtonClick() {
        try {
            const snapshot = this.getShareableStateSnapshot();
            const payload = await this.encodeShareState(snapshot);
            const shareUrl = `${window.location.origin}${window.location.pathname}${window.location.search}#${SHARE_HASH_KEY}=${payload}`;

            if (shareUrl.length > MAX_SHARE_URL_LENGTH) {
                await this.showAlertModal('This diagram is too large to share as a URL.');
                return;
            }

            if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
                await navigator.clipboard.writeText(shareUrl);
                await this.showAlertModal('Share link copied to clipboard.');
                return;
            }

            await this.showPromptModal('Copy this share URL', shareUrl);
        } catch (error) {
            console.error('Failed to create share URL:', error);
            await this.showAlertModal('Unable to generate share URL.');
        }
    }

    showPromptModal(message, defaultValue = '', options = {}) {
        return new Promise((resolve) => {
            const overlay = document.getElementById('custom-modal-overlay');
            const msgEl   = document.getElementById('custom-modal-message');
            const input   = document.getElementById('custom-modal-input');
            const ratioFields = document.getElementById('custom-modal-ratio-fields');
            const symbols = document.getElementById('custom-modal-symbols');
            const errorEl = document.getElementById('custom-modal-error');
            const confirm = document.getElementById('custom-modal-confirm');
            const cancel  = document.getElementById('custom-modal-cancel');
            const symbolByKey = {
                alpha: '\u03B1',
                beta: '\u03B2',
                gamma: '\u03B3',
                delta: '\u03B4',
                theta: '\u03B8',
                phi: '\u03C6',
                degree: '\u00B0'
            };
            const allowQuickSymbols = options.quickSymbols === true;

            msgEl.textContent = message;
            input.value = defaultValue;
            input.hidden = false;
            errorEl.textContent = '';
            if (ratioFields) {
                ratioFields.hidden = true;
            }
            if (symbols) {
                symbols.hidden = !allowQuickSymbols;
            }
            overlay.classList.add('show');
            overlay.setAttribute('aria-hidden', 'false');

            setTimeout(() => { input.focus(); input.select(); }, 50);

            const close = (value) => {
                overlay.classList.remove('show');
                overlay.setAttribute('aria-hidden', 'true');
                confirm.removeEventListener('click', onConfirm);
                cancel.removeEventListener('click', onCancel);
                overlay.removeEventListener('click', onBackdrop);
                input.removeEventListener('keydown', onKey);
                input.hidden = false;
                if (ratioFields) {
                    ratioFields.hidden = true;
                }
                if (symbols) {
                    symbols.removeEventListener('click', onSymbolClick);
                    symbols.hidden = true;
                }
                resolve(value);
            };

            const onConfirm = () => close(input.value);
            const onCancel  = () => close(null);
            const onBackdrop = (e) => { if (e.target === overlay) close(null); };
            const onSymbolClick = (e) => {
                const btn = e.target.closest('button[data-symbol-key]');
                if (!btn) {
                    return;
                }

                const symbol = symbolByKey[btn.dataset.symbolKey];
                if (!symbol) {
                    return;
                }

                const start = typeof input.selectionStart === 'number' ? input.selectionStart : input.value.length;
                const end = typeof input.selectionEnd === 'number' ? input.selectionEnd : input.value.length;
                input.focus();
                input.setRangeText(symbol, start, end, 'end');
            };
            const onKey = (e) => {
                if (e.key === 'Enter') { e.preventDefault(); close(input.value); }
                if (e.key === 'Escape') { e.preventDefault(); close(null); }
            };

            confirm.addEventListener('click', onConfirm);
            cancel.addEventListener('click', onCancel);
            overlay.addEventListener('click', onBackdrop);
            input.addEventListener('keydown', onKey);
            if (symbols && allowQuickSymbols) {
                symbols.addEventListener('click', onSymbolClick);
            }
        });
    }

    showRatioPromptModal(message, defaultLeft = 1, defaultRight = 2) {
        return new Promise((resolve) => {
            const overlay = document.getElementById('custom-modal-overlay');
            const msgEl = document.getElementById('custom-modal-message');
            const input = document.getElementById('custom-modal-input');
            const ratioFields = document.getElementById('custom-modal-ratio-fields');
            const leftInput = document.getElementById('custom-modal-ratio-left');
            const rightInput = document.getElementById('custom-modal-ratio-right');
            const symbols = document.getElementById('custom-modal-symbols');
            const errorEl = document.getElementById('custom-modal-error');
            const confirm = document.getElementById('custom-modal-confirm');
            const cancel = document.getElementById('custom-modal-cancel');

            if (!overlay || !msgEl || !input || !ratioFields || !leftInput || !rightInput || !errorEl || !confirm || !cancel) {
                resolve(null);
                return;
            }

            msgEl.textContent = message;
            input.hidden = true;
            ratioFields.hidden = false;
            leftInput.value = `${defaultLeft}`;
            rightInput.value = `${defaultRight}`;
            errorEl.textContent = '';
            if (symbols) {
                symbols.hidden = true;
            }
            overlay.classList.add('show');
            overlay.setAttribute('aria-hidden', 'false');

            setTimeout(() => {
                leftInput.focus();
                leftInput.select();
            }, 50);

            const close = (value) => {
                overlay.classList.remove('show');
                overlay.setAttribute('aria-hidden', 'true');
                input.hidden = false;
                ratioFields.hidden = true;
                confirm.removeEventListener('click', onConfirm);
                cancel.removeEventListener('click', onCancel);
                overlay.removeEventListener('click', onBackdrop);
                leftInput.removeEventListener('keydown', onKey);
                rightInput.removeEventListener('keydown', onKey);
                resolve(value);
            };

            const onConfirm = () => {
                const ratio = this.reduceRatio(leftInput.value, rightInput.value);
                if (!ratio) {
                    errorEl.textContent = 'Enter positive whole numbers for both sides of the ratio.';
                    return;
                }

                close(ratio);
            };
            const onCancel = () => close(null);
            const onBackdrop = (e) => { if (e.target === overlay) close(null); };
            const onKey = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    onConfirm();
                }
                if (e.key === 'Escape') {
                    e.preventDefault();
                    close(null);
                }
            };

            confirm.addEventListener('click', onConfirm);
            cancel.addEventListener('click', onCancel);
            overlay.addEventListener('click', onBackdrop);
            leftInput.addEventListener('keydown', onKey);
            rightInput.addEventListener('keydown', onKey);
        });
    }

    showAlertModal(message) {
        return new Promise((resolve) => {
            const overlay = document.getElementById('custom-modal-overlay');
            const msgEl   = document.getElementById('custom-modal-message');
            const input   = document.getElementById('custom-modal-input');
            const symbols = document.getElementById('custom-modal-symbols');
            const errorEl = document.getElementById('custom-modal-error');
            const confirm = document.getElementById('custom-modal-confirm');
            const cancel  = document.getElementById('custom-modal-cancel');

            msgEl.textContent = message;
            input.style.display = 'none';
            if (symbols) {
                symbols.hidden = true;
            }
            errorEl.textContent = '';
            cancel.style.display = 'none';
            overlay.classList.add('show');
            overlay.setAttribute('aria-hidden', 'false');

            const close = () => {
                overlay.classList.remove('show');
                overlay.setAttribute('aria-hidden', 'true');
                input.style.display = '';
                cancel.style.display = '';
                confirm.removeEventListener('click', close);
                overlay.removeEventListener('keydown', onKey);
                overlay.removeEventListener('click', onBackdrop);
                resolve();
            };

            const onKey = (e) => { if (e.key === 'Enter' || e.key === 'Escape') { e.preventDefault(); close(); } };
            const onBackdrop = (e) => { if (e.target === overlay) close(); };

            // Defer focus and listeners past the current key-event cycle so that
            // an Enter keyup from the prompt modal cannot immediately dismiss this alert.
            setTimeout(() => {
                confirm.focus();
                confirm.addEventListener('click', close);
                overlay.addEventListener('keydown', onKey);
                overlay.addEventListener('click', onBackdrop);
            }, 150);
        });
    }

    showConfirmModal(message, options = {}) {
        return new Promise((resolve) => {
            const overlay = document.getElementById('custom-modal-overlay');
            const msgEl = document.getElementById('custom-modal-message');
            const input = document.getElementById('custom-modal-input');
            const ratioFields = document.getElementById('custom-modal-ratio-fields');
            const symbols = document.getElementById('custom-modal-symbols');
            const errorEl = document.getElementById('custom-modal-error');
            const confirm = document.getElementById('custom-modal-confirm');
            const cancel = document.getElementById('custom-modal-cancel');

            if (!overlay || !msgEl || !input || !confirm || !cancel) {
                resolve(false);
                return;
            }

            const originalConfirmText = confirm.textContent;
            const originalCancelText = cancel.textContent;

            msgEl.textContent = message;
            input.hidden = true;
            if (ratioFields) {
                ratioFields.hidden = true;
            }
            if (symbols) {
                symbols.hidden = true;
            }
            errorEl.textContent = '';
            confirm.textContent = options.confirmText || 'OK';
            cancel.textContent = options.cancelText || 'Cancel';
            cancel.style.display = '';

            overlay.classList.add('show');
            overlay.setAttribute('aria-hidden', 'false');

            const close = (value) => {
                overlay.classList.remove('show');
                overlay.setAttribute('aria-hidden', 'true');
                input.hidden = false;
                if (ratioFields) {
                    ratioFields.hidden = true;
                }
                if (symbols) {
                    symbols.hidden = true;
                }
                confirm.textContent = originalConfirmText;
                cancel.textContent = originalCancelText;
                confirm.removeEventListener('click', onConfirm);
                cancel.removeEventListener('click', onCancel);
                overlay.removeEventListener('click', onBackdrop);
                overlay.removeEventListener('keydown', onKey);
                resolve(value);
            };

            const onConfirm = () => close(true);
            const onCancel = () => close(false);
            const onBackdrop = (e) => { if (e.target === overlay) close(false); };
            const onKey = (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    close(true);
                }
                if (e.key === 'Escape') {
                    e.preventDefault();
                    close(false);
                }
            };

            setTimeout(() => {
                confirm.focus();
                confirm.addEventListener('click', onConfirm);
                cancel.addEventListener('click', onCancel);
                overlay.addEventListener('click', onBackdrop);
                overlay.addEventListener('keydown', onKey);
            }, 0);
        });
    }

    showToast(message, durationMs = 2200) {
        if (!message) return;

        if (this.toastHideTimer) {
            clearTimeout(this.toastHideTimer);
            this.toastHideTimer = null;
        }

        if (!this.toastEl) {
            const toast = document.createElement('div');
            toast.className = 'app-toast';
            toast.setAttribute('role', 'status');
            toast.setAttribute('aria-live', 'polite');
            document.body.appendChild(toast);
            this.toastEl = toast;
        }

        this.toastEl.textContent = message;
        this.toastEl.classList.add('show');

        this.toastHideTimer = setTimeout(() => {
            if (this.toastEl) {
                this.toastEl.classList.remove('show');
            }
            this.toastHideTimer = null;
        }, Math.max(800, durationMs));
    }

    setupCrashDiagnostics() {
        if (this._crashListenersBound) {
            return;
        }

        this.recordCrashEvent('app.init', {
            userAgent: navigator.userAgent,
            url: window.location.href,
            shortcut: this.crashReportShortcut,
            wasDiscardedAtLoad: this.wasDiscardedAtLoad,
            deviceMemoryGb: Number.isFinite(navigator.deviceMemory) ? navigator.deviceMemory : null,
            hardwareConcurrency: Number.isFinite(navigator.hardwareConcurrency) ? navigator.hardwareConcurrency : null
        });

        this._onCrashWindowError = (event) => {
            this.recordCrashEvent('window.error', {
                message: event?.message || 'Unknown error',
                source: event?.filename || '',
                line: event?.lineno || 0,
                column: event?.colno || 0,
                stack: event?.error?.stack || ''
            });
        };

        this._onCrashUnhandledRejection = (event) => {
            const reason = event?.reason;
            this.recordCrashEvent('window.unhandledrejection', {
                message: reason?.message || String(reason || 'Unknown rejection'),
                stack: reason?.stack || ''
            });
        };

        this._onCrashVisibilityChange = () => {
            this.recordCrashEvent('document.visibilitychange', {
                state: document.visibilityState
            });
            this.recordIntegrityCheck(`visibilitychange:${document.visibilityState}`);
        };

        this._onCrashPageHide = (event) => {
            this.recordCrashEvent('window.pagehide', {
                persisted: event?.persisted === true
            });
        };

        this._onCrashPageShow = (event) => {
            this.recordCrashEvent('window.pageshow', {
                persisted: event?.persisted === true
            });
            this.recordIntegrityCheck('pageshow');
        };

        this._onCrashFocus = () => {
            this.recordCrashEvent('window.focus', {});
            this.recordIntegrityCheck('focus');
        };

        this._onCrashBlur = () => {
            this.recordCrashEvent('window.blur', {});
        };

        this._onCrashWebglContextLost = (event) => {
            // Prevent default so the browser can attempt WebGL context restoration.
            event.preventDefault();
            this.recordCrashEvent('canvas.webglcontextlost', {});
        };

        this._onCrashWebglContextRestored = () => {
            this.recordCrashEvent('canvas.webglcontextrestored', {});
            this.recordIntegrityCheck('webglcontextrestored');
        };

        this._onCrashFreeze = () => {
            this.recordCrashEvent('document.freeze', {});
            this.recordIntegrityCheck('freeze');
        };

        this._onCrashResume = () => {
            this.recordCrashEvent('document.resume', {});
            this.recordIntegrityCheck('resume');
        };

        window.addEventListener('error', this._onCrashWindowError);
        window.addEventListener('unhandledrejection', this._onCrashUnhandledRejection);
        document.addEventListener('visibilitychange', this._onCrashVisibilityChange);
        window.addEventListener('pagehide', this._onCrashPageHide);
        window.addEventListener('pageshow', this._onCrashPageShow);
        window.addEventListener('focus', this._onCrashFocus);
        window.addEventListener('blur', this._onCrashBlur);
        document.addEventListener('freeze', this._onCrashFreeze);
        document.addEventListener('resume', this._onCrashResume);
        this.canvas?.addEventListener('webglcontextlost', this._onCrashWebglContextLost);
        this.canvas?.addEventListener('webglcontextrestored', this._onCrashWebglContextRestored);

        this.crashWatchdogIntervalId = window.setInterval(() => {
            if (document.visibilityState === 'hidden') {
                this.recordIntegrityCheck('interval:hidden');
            }
        }, this.crashWatchdogIntervalMs);

        this.recordIntegrityCheck('startup');

        this._crashListenersBound = true;
    }

    teardownCrashDiagnostics() {
        if (!this._crashListenersBound) {
            return;
        }

        window.removeEventListener('error', this._onCrashWindowError);
        window.removeEventListener('unhandledrejection', this._onCrashUnhandledRejection);
        document.removeEventListener('visibilitychange', this._onCrashVisibilityChange);
        window.removeEventListener('pagehide', this._onCrashPageHide);
        window.removeEventListener('pageshow', this._onCrashPageShow);
        window.removeEventListener('focus', this._onCrashFocus);
        window.removeEventListener('blur', this._onCrashBlur);
        document.removeEventListener('freeze', this._onCrashFreeze);
        document.removeEventListener('resume', this._onCrashResume);
        this.canvas?.removeEventListener('webglcontextlost', this._onCrashWebglContextLost);
        this.canvas?.removeEventListener('webglcontextrestored', this._onCrashWebglContextRestored);

        if (this.crashWatchdogIntervalId) {
            window.clearInterval(this.crashWatchdogIntervalId);
            this.crashWatchdogIntervalId = null;
        }

        this._crashListenersBound = false;
    }

    recordCrashEvent(type, payload = {}) {
        const entry = {
            timestamp: new Date().toISOString(),
            type,
            payload
        };

        this.crashReportEntries.push(entry);
        if (this.crashReportEntries.length > this.maxCrashReportEntries) {
            this.crashReportEntries.splice(0, this.crashReportEntries.length - this.maxCrashReportEntries);
        }
    }

    buildStateIntegritySnapshot(reason = 'manual') {
        const sectionCounts = {};
        Object.entries(this.objectSections || {}).forEach(([key, sec]) => {
            sectionCounts[key] = sec?.list?.children?.length ?? null;
        });

        const sceneObjectTypeCounts = {};
        (this.sceneObjects || []).forEach((entry) => {
            const type = entry?.type || 'unknown';
            sceneObjectTypeCounts[type] = (sceneObjectTypeCounts[type] || 0) + 1;
        });

        const stateShape = {
            reason,
            visibilityState: document.visibilityState,
            compositeSlots: this.compositeSlots?.length ?? 0,
            pointDefinitions: this.pointDefinitions?.length ?? 0,
            derivedPoints: this.derivedPoints?.length ?? 0,
            sceneObjects: this.sceneObjects?.length ?? 0,
            visibleSceneObjects: (this.sceneObjects || []).filter((entry) => entry.visible !== false).length,
            selectedPoints: this.selectedPoints?.length ?? 0,
            sceneChildren: this.scene?.children?.length ?? 0,
            pointMarkers: this.pointMarkers?.size ?? 0,
            pointSprites: this.pointSprites?.length ?? 0,
            labelSprites: this.labelSprites?.length ?? 0,
            objectSectionDomCounts: sectionCounts,
            sceneObjectTypeCounts,
            panelOpen: this.panelOpen,
            triangleExtractState: this.triangleExtractTransitionState,
            triangleExtractOpen: !!this.activeTriangleExtraction,
            addDropdownVisible: this.addDropdown?.style?.display === 'block',
            methodPresence: {
                addSlot: typeof this.addSlot === 'function',
                renderObjectsList: typeof this.renderObjectsList === 'function',
                buildComposite: typeof this.buildComposite === 'function'
            }
        };

        stateShape.stateSize = JSON.stringify(stateShape).length;
        return stateShape;
    }

    recordIntegrityCheck(reason = 'manual') {
        const snapshot = this.buildStateIntegritySnapshot(reason);
        const digest = [
            snapshot.visibilityState,
            snapshot.compositeSlots,
            snapshot.pointDefinitions,
            snapshot.derivedPoints,
            snapshot.sceneObjects,
            snapshot.visibleSceneObjects,
            snapshot.sceneChildren,
            snapshot.addDropdownVisible,
            snapshot.triangleExtractState
        ].join('|');

        if (reason === 'interval:hidden' && digest === this.lastIntegrityDigest) {
            return;
        }

        this.lastIntegrityDigest = digest;
        this.recordCrashEvent('state.integrity', snapshot);
    }

    isCrashReportOpen() {
        return !!this.crashReportOverlay?.classList.contains('show');
    }

    toggleCrashReport() {
        if (this.isCrashReportOpen()) {
            this.closeCrashReport();
            return;
        }

        this.openCrashReport();
    }

    openCrashReport() {
        if (!this.crashReportOverlay || !this.crashReportPre) {
            return;
        }

        this.crashReportOpenedAt = new Date().toISOString();
        this.recordIntegrityCheck('report:open');
        this.recordCrashEvent('crash-report.opened', {
            shortcut: this.crashReportShortcut
        });
        this.crashReportPre.textContent = this.buildCrashReportText();
        this.crashReportOverlay.classList.add('show');
        this.crashReportOverlay.setAttribute('aria-hidden', 'false');
        this.crashReportCloseBtn?.focus();
    }

    closeCrashReport() {
        if (!this.crashReportOverlay) {
            return;
        }

        this.crashReportOverlay.classList.remove('show');
        this.crashReportOverlay.setAttribute('aria-hidden', 'true');
    }

    refreshCrashReport() {
        if (!this.crashReportPre) {
            return;
        }

        this.recordIntegrityCheck('report:refresh');
        this.crashReportPre.textContent = this.buildCrashReportText();
    }

    buildCrashReportText() {
        const integritySnapshot = this.buildStateIntegritySnapshot('report:summary');
        const summary = {
            generatedAt: new Date().toISOString(),
            openedAt: this.crashReportOpenedAt,
            url: window.location.href,
            wasDiscardedAtLoad: this.wasDiscardedAtLoad,
            visibilityState: document.visibilityState,
            online: navigator.onLine,
            deviceMemoryGb: Number.isFinite(navigator.deviceMemory) ? navigator.deviceMemory : null,
            hardwareConcurrency: Number.isFinite(navigator.hardwareConcurrency) ? navigator.hardwareConcurrency : null,
            performanceMemory: performance?.memory ? {
                jsHeapSizeLimit: performance.memory.jsHeapSizeLimit,
                totalJSHeapSize: performance.memory.totalJSHeapSize,
                usedJSHeapSize: performance.memory.usedJSHeapSize
            } : null,
            panelOpen: this.panelOpen,
            activeCompositePrimitives: this.compositeSlots.length,
            points: this.getAllPoints().length,
            sceneObjects: this.sceneObjects.length,
            visibleSceneObjects: this.sceneObjects.filter((entry) => entry.visible !== false).length,
            selectedPoints: this.selectedPoints.length,
            triangleExtractState: this.triangleExtractTransitionState,
            triangleExtractOpen: this.isTriangleExtractionOpen(),
            addDropdownVisible: this.addDropdown?.style.display === 'block',
            rendererInfo: this.renderer?.info ? {
                geometries: this.renderer.info.memory?.geometries ?? null,
                textures: this.renderer.info.memory?.textures ?? null,
                calls: this.renderer.info.render?.calls ?? null,
                triangles: this.renderer.info.render?.triangles ?? null,
                points: this.renderer.info.render?.points ?? null,
                lines: this.renderer.info.render?.lines ?? null
            } : null,
            integritySnapshot
        };

        const lines = [];
        lines.push('Trimension Hidden Crash Report');
        lines.push(`Shortcut: ${this.crashReportShortcut}`);
        lines.push('');
        lines.push('[Summary]');
        lines.push(JSON.stringify(summary, null, 2));
        lines.push('');
        lines.push('[Recent Events]');

        if (this.crashReportEntries.length === 0) {
            lines.push('(no events recorded)');
        } else {
            this.crashReportEntries.slice(-80).forEach((entry) => {
                const payloadText = JSON.stringify(entry.payload || {});
                lines.push(`${entry.timestamp} | ${entry.type} | ${payloadText}`);
            });
        }

        return lines.join('\n');
    }

    async copyCrashReport() {
        if (!this.crashReportPre) {
            return;
        }

        const text = this.crashReportPre.textContent || this.buildCrashReportText();
        try {
            if (navigator.clipboard?.writeText) {
                await navigator.clipboard.writeText(text);
                this.showToast('Crash report copied.');
                return;
            }
        } catch {
            // Ignore and continue to fallback copy path.
        }

        const tempArea = document.createElement('textarea');
        tempArea.value = text;
        tempArea.setAttribute('readonly', 'true');
        tempArea.style.position = 'fixed';
        tempArea.style.left = '-9999px';
        document.body.appendChild(tempArea);
        tempArea.select();
        document.execCommand('copy');
        document.body.removeChild(tempArea);
        this.showToast('Crash report copied.');
    }

    isTriangleExtractionOpen() {
        return !!this.activeTriangleExtraction;
    }

    openTriangleExtraction(objectId) {
        if (this.triangleExtractTransitionState !== 'closed') {
            return;
        }

        if (this.isTriangleExtractionOpen()) {
            this.closeTriangleExtraction();
        }

        const item = this.sceneObjects.find((entry) => {
            if (entry.id !== objectId || !entry.definition || !Array.isArray(entry.definition.pointIds)) {
                return false;
            }

            return (entry.type === 'triangle' && entry.definition.pointIds.length === 3)
                || (entry.type === 'plane' && entry.definition.pointIds.length === 4);
        });
        if (!item) {
            return;
        }

        const points = item.definition.pointIds.map((pointId) => this.getPointById(pointId));
        if (points.some((point) => !point)) {
            return;
        }

        const worldPoints = points.map((point) => point.position.clone());
        const labels = points.map((point) => point.label || point.id);
        if (!this.triangleExtractOverlay || !this.triangleExtractModal || !this.triangleExtractFlightSvg) {
            return;
        }
        const flightColor = this.getTriangleExtractionColor(item.definition?.color);

        const sourcePoints = worldPoints.map((point) => this.projectWorldPointToViewport(point));
        const isTriangleExtraction = item.type === 'triangle';

        let layout = item.type === 'plane'
            ? this.buildCameraAwarePlaneExtractionLayout(
                worldPoints,
                sourcePoints,
                this.getTriangleExtractionStageAspectRatio()
            )
            : this.buildCameraAwareTriangleExtractionLayout(
                worldPoints,
                sourcePoints,
                this.getTriangleExtractionStageAspectRatio()
            );

        if (isTriangleExtraction && layout) {
            const cameraAware = this.applyCameraAwareTriangleOrientation(layout, sourcePoints);
            layout = cameraAware.layout;
        }

        if (!layout) {
            return;
        }

        this.activeTriangleExtraction = {
            objectId: item.id,
            type: item.type,
            pointIds: [...item.definition.pointIds],
            layout,
            baseLayout: layout,
            labels: [...labels],
            item,
            color: flightColor,
            orientationQuarterTurns: 0,
            orientationFlipped: false
        };
        this.lastFocusedElementBeforeTriangleExtract = document.activeElement instanceof HTMLElement ? document.activeElement : null;
        this.controls.enabled = false;
        this._keysHeld?.clear();
        this.triangleExtractTransitionState = 'opening';

        this.triangleExtractOverlay.classList.add('show', 'pre-open');
        this.triangleExtractOverlay.classList.remove('settled');
        this.triangleExtractOverlay.setAttribute('aria-hidden', 'false');

        this.populateTriangleExtractionModal(layout, labels, item, flightColor);
        this.updateTriangleExtractionOrientationButtons();

        const animationDurationMs = 1450;

        requestAnimationFrame(() => {
            if (!this.activeTriangleExtraction) {
                return;
            }

            this.triangleExtractOverlay.classList.remove('pre-open');
            const destinationPoints = this.getTriangleExtractionDestinationPoints(layout);
            this.updateTriangleExtractFlightStyle(flightColor);

            if (!destinationPoints) {
                this.finalizeTriangleExtractionReveal();
                return;
            }

            if (animationDurationMs === 0) {
                this.renderTriangleExtractFlight(destinationPoints);
                this.finalizeTriangleExtractionReveal();
                return;
            }

            const startTime = performance.now();
            const step = (now) => {
                if (!this.activeTriangleExtraction) {
                    this.triangleExtractAnimationFrame = null;
                    return;
                }

                const rawProgress = Math.min(1, (now - startTime) / animationDurationMs);
                const eased = 1 - Math.pow(1 - rawProgress, 3);
                const currentPoints = sourcePoints.map((point, index) => ({
                    x: THREE.MathUtils.lerp(point.x, destinationPoints[index].x, eased),
                    y: THREE.MathUtils.lerp(point.y, destinationPoints[index].y, eased)
                }));
                this.renderTriangleExtractFlight(currentPoints);

                if (rawProgress >= 1) {
                    this.triangleExtractAnimationFrame = null;
                    this.finalizeTriangleExtractionReveal();
                    return;
                }

                this.triangleExtractAnimationFrame = window.requestAnimationFrame(step);
            };

            this.renderTriangleExtractFlight(sourcePoints);
            this.triangleExtractAnimationFrame = window.requestAnimationFrame(step);
        });
    }

    getTriangleSignedArea2D(points) {
        if (!Array.isArray(points) || points.length < 3) {
            return 0;
        }

        const a = points[0];
        const b = points[1];
        const c = points[2];
        return ((b.x - a.x) * (c.y - a.y)) - ((b.y - a.y) * (c.x - a.x));
    }

    scoreTriangleLayoutAgainstCamera(layout, cameraPoints) {
        if (!layout?.points2D || layout.points2D.length < 3 || !Array.isArray(cameraPoints) || cameraPoints.length < 3) {
            return -Infinity;
        }

        const normalizedLayout = this.normalizePointsForComparison(layout.points2D);
        const normalizedCamera = this.normalizePointsForComparison(cameraPoints);
        if (!normalizedLayout || !normalizedCamera) {
            return -Infinity;
        }

        // Vertex-to-vertex fit after translation/scale normalization.
        let pointError = 0;
        for (let index = 0; index < 3; index += 1) {
            const dx = normalizedLayout[index].x - normalizedCamera[index].x;
            const dy = normalizedLayout[index].y - normalizedCamera[index].y;
            pointError += (dx * dx) + (dy * dy);
        }

        // Edge-direction agreement stabilizes ties where point fit is very close.
        const edgePairs = [[0, 1], [1, 2], [2, 0]];
        let edgeDirectionScore = 0;
        edgePairs.forEach(([startIndex, endIndex]) => {
            const layoutDx = normalizedLayout[endIndex].x - normalizedLayout[startIndex].x;
            const layoutDy = normalizedLayout[endIndex].y - normalizedLayout[startIndex].y;
            const cameraDx = normalizedCamera[endIndex].x - normalizedCamera[startIndex].x;
            const cameraDy = normalizedCamera[endIndex].y - normalizedCamera[startIndex].y;

            const layoutLen = Math.hypot(layoutDx, layoutDy);
            const cameraLen = Math.hypot(cameraDx, cameraDy);
            if (layoutLen <= 1e-6 || cameraLen <= 1e-6) {
                return;
            }

            edgeDirectionScore += ((layoutDx / layoutLen) * (cameraDx / cameraLen))
                + ((layoutDy / layoutLen) * (cameraDy / cameraLen));
        });

        return edgeDirectionScore - (pointError * 2.5);
    }

    scorePolygonLayoutAgainstCamera(layout, cameraPoints) {
        if (!layout?.points2D || layout.points2D.length < 3 || !Array.isArray(cameraPoints) || cameraPoints.length !== layout.points2D.length) {
            return -Infinity;
        }

        const normalizedLayout = this.normalizePointsForComparison(layout.points2D);
        const normalizedCamera = this.normalizePointsForComparison(cameraPoints);
        if (!normalizedLayout || !normalizedCamera) {
            return -Infinity;
        }

        let pointError = 0;
        for (let index = 0; index < normalizedLayout.length; index += 1) {
            const dx = normalizedLayout[index].x - normalizedCamera[index].x;
            const dy = normalizedLayout[index].y - normalizedCamera[index].y;
            pointError += (dx * dx) + (dy * dy);
        }

        let edgeDirectionScore = 0;
        for (let index = 0; index < normalizedLayout.length; index += 1) {
            const nextIndex = (index + 1) % normalizedLayout.length;
            const layoutDx = normalizedLayout[nextIndex].x - normalizedLayout[index].x;
            const layoutDy = normalizedLayout[nextIndex].y - normalizedLayout[index].y;
            const cameraDx = normalizedCamera[nextIndex].x - normalizedCamera[index].x;
            const cameraDy = normalizedCamera[nextIndex].y - normalizedCamera[index].y;

            const layoutLen = Math.hypot(layoutDx, layoutDy);
            const cameraLen = Math.hypot(cameraDx, cameraDy);
            if (layoutLen <= 1e-6 || cameraLen <= 1e-6) {
                continue;
            }

            edgeDirectionScore += ((layoutDx / layoutLen) * (cameraDx / cameraLen))
                + ((layoutDy / layoutLen) * (cameraDy / cameraLen));
        }

        return edgeDirectionScore - (pointError * 2.5);
    }

    normalizePointsForComparison(points) {
        if (!Array.isArray(points) || points.length < 3) {
            return null;
        }

        const centroid = points.reduce((acc, point) => ({
            x: acc.x + point.x,
            y: acc.y + point.y
        }), { x: 0, y: 0 });
        centroid.x /= points.length;
        centroid.y /= points.length;

        const centered = points.map((point) => ({
            x: point.x - centroid.x,
            y: point.y - centroid.y
        }));

        const rms = Math.sqrt(
            centered.reduce((acc, point) => acc + (point.x * point.x) + (point.y * point.y), 0)
            / centered.length
        );
        if (rms <= 1e-6) {
            return null;
        }

        return centered.map((point) => ({
            x: point.x / rms,
            y: point.y / rms
        }));
    }

    applyCameraAwarePolygonOrientation(layout, cameraPoints) {
        if (!layout?.points2D || layout.points2D.length < 3) {
            return {
                layout,
                quarterTurns: 0,
                flipped: false,
                score: -Infinity
            };
        }

        const candidates = [];
        for (let quarterTurns = 0; quarterTurns < 4; quarterTurns += 1) {
            for (let flipIndex = 0; flipIndex < 2; flipIndex += 1) {
                const isFlipped = flipIndex === 1;
                const candidateLayout = (quarterTurns === 0 && !isFlipped)
                    ? layout
                    : this.buildTransformedExtractionLayout(layout, quarterTurns, isFlipped);

                if (!candidateLayout) {
                    continue;
                }

                const cameraScore = this.scorePolygonLayoutAgainstCamera(candidateLayout, cameraPoints);
                candidates.push({
                    layout: candidateLayout,
                    quarterTurns,
                    flipped: isFlipped,
                    cameraScore,
                    score: cameraScore
                });
            }
        }

        if (candidates.length === 0) {
            return {
                layout,
                quarterTurns: 0,
                flipped: false,
                score: -Infinity
            };
        }

        const selected = candidates.reduce((best, candidate) => {
            if (!best || candidate.cameraScore > best.cameraScore + 1e-6) {
                return candidate;
            }

            if (!best || Math.abs(candidate.cameraScore - best.cameraScore) <= 1e-6) {
                if (candidate.quarterTurns < best.quarterTurns) {
                    return candidate;
                }
                if (candidate.quarterTurns === best.quarterTurns && Number(candidate.flipped) < Number(best.flipped)) {
                    return candidate;
                }
            }

            return best;
        }, null);

        return selected || {
            layout,
            quarterTurns: 0,
            flipped: false,
            score: -Infinity
        };
    }

    applyCameraAwareTriangleOrientation(layout, cameraPoints) {
        if (!layout?.points2D || layout.points2D.length < 3) {
            return {
                layout,
                quarterTurns: 0,
                flipped: false
            };
        }

        const candidates = [];
        for (let quarterTurns = 0; quarterTurns < 4; quarterTurns += 1) {
            for (let flipIndex = 0; flipIndex < 2; flipIndex += 1) {
                const isFlipped = flipIndex === 1;
                const candidateLayout = (quarterTurns === 0 && !isFlipped)
                    ? layout
                    : this.buildTransformedExtractionLayout(layout, quarterTurns, isFlipped);

                if (!candidateLayout) {
                    continue;
                }

                const cameraScore = this.scoreTriangleLayoutAgainstCamera(candidateLayout, cameraPoints);
                candidates.push({
                    layout: candidateLayout,
                    quarterTurns,
                    flipped: isFlipped,
                    cameraScore,
                    score: cameraScore
                });
            }
        }

        if (candidates.length === 0) {
            return {
                layout,
                quarterTurns: 0,
                flipped: false,
                score: -Infinity
            };
        }

        const selected = candidates.reduce((best, candidate) => {
            if (!best || candidate.cameraScore > best.cameraScore + 1e-6) {
                return candidate;
            }

            if (!best || Math.abs(candidate.cameraScore - best.cameraScore) <= 1e-6) {
                if (candidate.quarterTurns < best.quarterTurns) {
                    return candidate;
                }
                if (candidate.quarterTurns === best.quarterTurns && Number(candidate.flipped) < Number(best.flipped)) {
                    return candidate;
                }
            }

            return best;
        }, null);

        return selected || {
            layout,
            quarterTurns: 0,
            flipped: false,
            score: -Infinity
        };
    }

    buildTriangleExtractionCandidateLayouts(worldPoints, targetAspectRatio = 1000 / 760) {
        if (!Array.isArray(worldPoints) || worldPoints.length !== 3) {
            return [];
        }

        const sideAB = worldPoints[0].distanceTo(worldPoints[1]);
        const sideBC = worldPoints[1].distanceTo(worldPoints[2]);
        const sideCA = worldPoints[2].distanceTo(worldPoints[0]);
        const baseLength = sideAB;
        if (baseLength <= 1e-6) {
            return [];
        }

        const cX = (sideCA * sideCA - sideBC * sideBC + baseLength * baseLength) / (2 * baseLength);
        const cY = Math.sqrt(Math.max(0, sideCA * sideCA - cX * cX));
        if (cY <= 1e-6) {
            return [];
        }

        const rawPoints = [
            { x: 0, y: 0 },
            { x: baseLength, y: 0 },
            { x: cX, y: cY }
        ];

        const stageWidth = 1000;
        const safeAspect = Number.isFinite(targetAspectRatio) && targetAspectRatio > 0.1
            ? targetAspectRatio
            : (1000 / 760);
        const stageHeight = Math.max(1, Math.round(stageWidth / safeAspect));
        const padding = this.getExtractionLayoutPadding(stageWidth, stageHeight);

        return this.getTriangleOrientationCandidates(rawPoints)
            .map((candidate) => {
                const minX = Math.min(...candidate.points.map((point) => point.x));
                const maxX = Math.max(...candidate.points.map((point) => point.x));
                const minY = Math.min(...candidate.points.map((point) => point.y));
                const maxY = Math.max(...candidate.points.map((point) => point.y));
                const width = Math.max(maxX - minX, 1e-6);
                const height = Math.max(maxY - minY, 1e-6);
                const scaleScore = Math.min((stageWidth - padding.x * 2) / width, (stageHeight - padding.y * 2) / height);
                const points2D = this.fitExtractionPointsToStage(candidate.points, stageWidth, stageHeight, padding.x, padding.y);
                const layout = points2D
                    ? this.buildExtractionLayoutFromPoints(points2D, stageWidth, stageHeight)
                    : null;

                return layout
                    ? {
                        layout,
                        scaleScore,
                        startIndex: candidate.startIndex,
                        endIndex: candidate.endIndex,
                        apexIndex: candidate.apexIndex
                    }
                    : null;
            })
            .filter(Boolean);
    }

    buildCameraAwareTriangleExtractionLayout(worldPoints, cameraPoints, targetAspectRatio = 1000 / 760) {
        const candidateLayouts = this.buildTriangleExtractionCandidateLayouts(worldPoints, targetAspectRatio);
        if (candidateLayouts.length === 0) {
            return this.buildTriangleExtractionLayout(worldPoints, targetAspectRatio);
        }

        let bestCandidate = null;
        candidateLayouts.forEach((candidate) => {
            const oriented = this.applyCameraAwareTriangleOrientation(candidate.layout, cameraPoints);
            const totalScore = (Number.isFinite(oriented.score) ? oriented.score : -Infinity)
                + (candidate.scaleScore * 0.015);

            if (!bestCandidate || totalScore > bestCandidate.totalScore + 1e-6) {
                bestCandidate = {
                    layout: oriented.layout,
                    totalScore
                };
            }
        });

        return bestCandidate?.layout || this.buildTriangleExtractionLayout(worldPoints, targetAspectRatio);
    }

    buildPolygonExtractionCandidateLayouts(rawPoints, targetAspectRatio = 1000 / 760) {
        if (!Array.isArray(rawPoints) || rawPoints.length < 3) {
            return [];
        }

        const stageWidth = 1000;
        const safeAspect = Number.isFinite(targetAspectRatio) && targetAspectRatio > 0.1
            ? targetAspectRatio
            : (1000 / 760);
        const stageHeight = Math.max(1, Math.round(stageWidth / safeAspect));
        const padding = this.getExtractionLayoutPadding(stageWidth, stageHeight);

        return this.getPolygonOrientationCandidates(rawPoints)
            .map((candidate) => {
                const minX = Math.min(...candidate.points.map((point) => point.x));
                const maxX = Math.max(...candidate.points.map((point) => point.x));
                const minY = Math.min(...candidate.points.map((point) => point.y));
                const maxY = Math.max(...candidate.points.map((point) => point.y));
                const width = Math.max(maxX - minX, 1e-6);
                const height = Math.max(maxY - minY, 1e-6);
                const scaleScore = Math.min((stageWidth - padding.x * 2) / width, (stageHeight - padding.y * 2) / height);
                const points2D = this.fitExtractionPointsToStage(candidate.points, stageWidth, stageHeight, padding.x, padding.y);
                const layout = points2D
                    ? this.buildExtractionLayoutFromPoints(points2D, stageWidth, stageHeight)
                    : null;

                return layout
                    ? {
                        layout,
                        scaleScore
                    }
                    : null;
            })
            .filter(Boolean);
    }

    buildCameraAwarePolygonExtractionLayout(rawPoints, cameraPoints, targetAspectRatio = 1000 / 760) {
        const candidateLayouts = this.buildPolygonExtractionCandidateLayouts(rawPoints, targetAspectRatio);
        if (candidateLayouts.length === 0) {
            return this.buildPolygonExtractionLayout(rawPoints, targetAspectRatio);
        }

        let bestCandidate = null;
        candidateLayouts.forEach((candidate) => {
            const oriented = this.applyCameraAwarePolygonOrientation(candidate.layout, cameraPoints);
            const totalScore = (Number.isFinite(oriented.score) ? oriented.score : -Infinity)
                + (candidate.scaleScore * 0.015);

            if (!bestCandidate || totalScore > bestCandidate.totalScore + 1e-6) {
                bestCandidate = {
                    layout: oriented.layout,
                    totalScore
                };
            }
        });

        return bestCandidate?.layout || this.buildPolygonExtractionLayout(rawPoints, targetAspectRatio);
    }

    buildCameraAwarePlaneExtractionLayout(worldPoints, cameraPoints, targetAspectRatio = 1000 / 760) {
        if (!Array.isArray(worldPoints) || worldPoints.length !== 4) {
            return null;
        }

        const origin = worldPoints[0];
        const uVector = worldPoints[1].clone().sub(origin);
        if (uVector.lengthSq() <= 1e-12) {
            return null;
        }

        let normal = null;
        for (let index = 2; index < worldPoints.length; index += 1) {
            const candidate = new THREE.Vector3().crossVectors(uVector, worldPoints[index].clone().sub(origin));
            if (candidate.lengthSq() > 1e-12) {
                normal = candidate.normalize();
                break;
            }
        }

        if (!normal) {
            return null;
        }

        const u = uVector.clone().normalize();
        const v = new THREE.Vector3().crossVectors(normal, u).normalize();
        const rawPoints = worldPoints.map((point) => {
            const relative = point.clone().sub(origin);
            return {
                x: relative.dot(u),
                y: relative.dot(v)
            };
        });

        return this.buildCameraAwarePolygonExtractionLayout(rawPoints, cameraPoints, targetAspectRatio);
    }

    closeTriangleExtraction(options = {}) {
        if (!this.activeTriangleExtraction || !this.triangleExtractOverlay || !this.triangleExtractModal) {
            return;
        }

        const forceClose = options?.force === true;
        if (!forceClose && this.triangleExtractTransitionState !== 'open') {
            return;
        }

        if (forceClose && this.triangleExtractTransitionState !== 'open') {
            this.finishTriangleExtractionClose();
            return;
        }

        if (this.triangleExtractSettleTimer) {
            window.clearTimeout(this.triangleExtractSettleTimer);
            this.triangleExtractSettleTimer = null;
        }
        if (this.triangleExtractAnimationFrame) {
            window.cancelAnimationFrame(this.triangleExtractAnimationFrame);
            this.triangleExtractAnimationFrame = null;
        }

        const extraction = this.activeTriangleExtraction;
        const animationDurationMs = 1450;
        const points = (extraction.pointIds || []).map((pointId) => this.getPointById(pointId));
        const worldPoints = points.some((point) => !point)
            ? null
            : points.map((point) => point.position.clone());
        const returnPoints = worldPoints
            ? worldPoints.map((point) => this.projectWorldPointToViewport(point))
            : null;
        const startPoints = extraction.layout
            ? this.getTriangleExtractionDestinationPoints(extraction.layout)
            : null;

        // Hide modal chrome/text first, then run a reverse flight animation.
        this.triangleExtractTransitionState = 'closing';
        this.triangleExtractOverlay.classList.remove('settled', 'pre-open');
        if (extraction.color) {
            this.updateTriangleExtractFlightStyle(extraction.color);
        }

        if (!startPoints || !returnPoints || animationDurationMs === 0) {
            this.finishTriangleExtractionClose();
            return;
        }

        this.renderTriangleExtractFlight(startPoints);
        const startTime = performance.now();
        const step = (now) => {
            if (!this.activeTriangleExtraction) {
                this.triangleExtractAnimationFrame = null;
                return;
            }

            const rawProgress = Math.min(1, (now - startTime) / animationDurationMs);
            const eased = 1 - Math.pow(1 - rawProgress, 3);
            const currentPoints = startPoints.map((point, index) => ({
                x: THREE.MathUtils.lerp(point.x, returnPoints[index].x, eased),
                y: THREE.MathUtils.lerp(point.y, returnPoints[index].y, eased)
            }));
            this.renderTriangleExtractFlight(currentPoints);

            if (rawProgress >= 1) {
                this.triangleExtractAnimationFrame = null;
                this.finishTriangleExtractionClose();
                return;
            }

            this.triangleExtractAnimationFrame = window.requestAnimationFrame(step);
        };

        this.triangleExtractAnimationFrame = window.requestAnimationFrame(step);
    }

    finishTriangleExtractionClose() {
        this.activeTriangleExtraction = null;
        this.triangleExtractTransitionState = 'closed';
        this.triangleExtractOverlay.classList.remove('show', 'pre-open', 'settled');
        this.triangleExtractOverlay.setAttribute('aria-hidden', 'true');
        this.controls.enabled = true;
        this.clearTriangleExtractFlight();
        this.updateTriangleExtractionOrientationButtons();

        if (this.lastFocusedElementBeforeTriangleExtract?.focus) {
            this.lastFocusedElementBeforeTriangleExtract.focus();
        }
        this.lastFocusedElementBeforeTriangleExtract = null;
    }

    projectWorldPointToViewport(worldPoint) {
        const rect = this.canvas.getBoundingClientRect();
        const projected = worldPoint.clone().project(this.camera);
        return {
            x: rect.left + ((projected.x + 1) / 2) * rect.width,
            y: rect.top + ((1 - projected.y) / 2) * rect.height
        };
    }

    getTriangleExtractionDestinationPoints(layout) {
        if (!this.triangleExtractSvg) {
            return null;
        }

        const rect = this.triangleExtractSvg.getBoundingClientRect();
        if (!rect.width || !rect.height) {
            return null;
        }

        return layout.points2D.map((point) => ({
            x: rect.left + (point.x / layout.stageWidth) * rect.width,
            y: rect.top + (point.y / layout.stageHeight) * rect.height
        }));
    }

    getTriangleExtractionColor(colorValue) {
        const numericColor = Number.isFinite(colorValue) ? colorValue : 0xff595e;
        const hex = `#${numericColor.toString(16).padStart(6, '0')}`;
        return {
            stroke: hex,
            fill: `${hex}55`
        };
    }

    updateTriangleExtractFlightStyle(color) {
        if (this.triangleExtractFlightPolygon) {
            this.triangleExtractFlightPolygon.style.fill = color.fill;
        }
        if (this.triangleExtractFlightOutline) {
            this.triangleExtractFlightOutline.style.stroke = color.stroke;
        }
    }

    renderTriangleExtractFlight(points) {
        if (!this.triangleExtractFlightSvg || !this.triangleExtractFlightPolygon || !this.triangleExtractFlightOutline || points.length < 3) {
            return;
        }

        const overlayRect = this.triangleExtractOverlay.getBoundingClientRect();
        this.triangleExtractFlightSvg.setAttribute('viewBox', `0 0 ${Math.max(1, overlayRect.width)} ${Math.max(1, overlayRect.height)}`);
        const localPoints = points.map((point) => ({
            x: point.x - overlayRect.left,
            y: point.y - overlayRect.top
        }));
        const pointsString = localPoints.map((point) => `${point.x},${point.y}`).join(' ');
        this.triangleExtractFlightPolygon.setAttribute('points', pointsString);
        this.triangleExtractFlightOutline.setAttribute('points', `${pointsString} ${localPoints[0].x},${localPoints[0].y}`);
    }

    clearTriangleExtractFlight() {
        this.triangleExtractFlightPolygon?.setAttribute('points', '');
        this.triangleExtractFlightOutline?.setAttribute('points', '');
    }

    finalizeTriangleExtractionReveal() {
        this.triangleExtractSettleTimer = window.setTimeout(() => {
            if (!this.activeTriangleExtraction || !this.triangleExtractOverlay) {
                return;
            }

            this.clearTriangleExtractFlight();
            this.triangleExtractOverlay.classList.add('settled');
            this.triangleExtractTransitionState = 'open';
            this.updateTriangleExtractionOrientationButtons();
            this.triangleExtractCloseBtn?.focus();
            this.triangleExtractSettleTimer = null;
        }, 30);
    }

    getTriangleExtractionStageAspectRatio() {
        const rect = this.triangleExtractStage?.getBoundingClientRect();
        if (rect && rect.width > 1 && rect.height > 1) {
            return rect.width / rect.height;
        }

        const modalRect = this.triangleExtractModal?.getBoundingClientRect();
        if (modalRect && modalRect.width > 1 && modalRect.height > 1) {
            // Approximate content area by reserving room for the header.
            const estimatedStageHeight = Math.max(1, modalRect.height - 84);
            return modalRect.width / estimatedStageHeight;
        }

        return window.innerWidth > 1 && window.innerHeight > 1
            ? (window.innerWidth / window.innerHeight)
            : (1000 / 760);
    }

    getTriangleOrientationCandidates(rawPoints) {
        const candidates = [];
        const baseDefinitions = [
            { startIndex: 0, endIndex: 1, apexIndex: 2 },
            { startIndex: 1, endIndex: 2, apexIndex: 0 },
            { startIndex: 2, endIndex: 0, apexIndex: 1 }
        ];

        baseDefinitions.forEach((base) => {
            const start = rawPoints[base.startIndex];
            const end = rawPoints[base.endIndex];
            const dx = end.x - start.x;
            const dy = end.y - start.y;
            const baseLength = Math.hypot(dx, dy);
            if (baseLength <= 1e-6) {
                return;
            }

            const ux = dx / baseLength;
            const uy = dy / baseLength;
            const oriented = rawPoints.map((point) => {
                const relX = point.x - start.x;
                const relY = point.y - start.y;
                return {
                    x: relX * ux + relY * uy,
                    y: -relX * uy + relY * ux
                };
            });

            if (oriented[base.apexIndex].y < 0) {
                oriented.forEach((point) => {
                    point.y *= -1;
                });
            }

            candidates.push({
                points: oriented,
                baseLength,
                startIndex: base.startIndex,
                endIndex: base.endIndex,
                apexIndex: base.apexIndex
            });
        });

        return candidates;
    }

    getPolygonOrientationCandidates(rawPoints) {
        const candidates = [];

        for (let index = 0; index < rawPoints.length; index += 1) {
            const start = rawPoints[index];
            const end = rawPoints[(index + 1) % rawPoints.length];
            const dx = end.x - start.x;
            const dy = end.y - start.y;
            const baseLength = Math.hypot(dx, dy);
            if (baseLength <= 1e-6) {
                continue;
            }

            const ux = dx / baseLength;
            const uy = dy / baseLength;
            const oriented = rawPoints.map((point) => {
                const relX = point.x - start.x;
                const relY = point.y - start.y;
                return {
                    x: relX * ux + relY * uy,
                    y: -relX * uy + relY * ux
                };
            });

            const signedArea = oriented.reduce((area, point, pointIndex) => {
                const nextPoint = oriented[(pointIndex + 1) % oriented.length];
                return area + (point.x * nextPoint.y) - (nextPoint.x * point.y);
            }, 0) / 2;

            if (signedArea < 0) {
                oriented.forEach((point) => {
                    point.y *= -1;
                });
            }

            candidates.push({
                points: oriented,
                baseLength
            });
        }

        return candidates;
    }

    getExtractionLayoutPadding(stageWidth, stageHeight) {
        return {
            x: Math.max(88, Math.min(128, stageWidth * 0.115)),
            y: Math.max(72, Math.min(104, stageHeight * 0.135))
        };
    }

    fitExtractionPointsToStage(rawPoints, stageWidth, stageHeight, paddingX = null, paddingY = null) {
        if (!Array.isArray(rawPoints) || rawPoints.length < 3) {
            return null;
        }

        const padding = this.getExtractionLayoutPadding(stageWidth, stageHeight);
        const safePaddingX = Number.isFinite(paddingX) ? paddingX : padding.x;
        const safePaddingY = Number.isFinite(paddingY) ? paddingY : padding.y;

        const minX = Math.min(...rawPoints.map((point) => point.x));
        const maxX = Math.max(...rawPoints.map((point) => point.x));
        const minY = Math.min(...rawPoints.map((point) => point.y));
        const maxY = Math.max(...rawPoints.map((point) => point.y));
        const width = Math.max(maxX - minX, 1e-6);
        const height = Math.max(maxY - minY, 1e-6);
        const scale = Math.min((stageWidth - safePaddingX * 2) / width, (stageHeight - safePaddingY * 2) / height);
        const offsetX = (stageWidth - width * scale) / 2 - minX * scale;
        const offsetY = (stageHeight - height * scale) / 2 + maxY * scale;

        return rawPoints.map((point) => ({
            x: offsetX + point.x * scale,
            y: offsetY - point.y * scale
        }));
    }

    buildExtractionLayoutFromPoints(points2D, stageWidth, stageHeight) {
        if (!Array.isArray(points2D) || points2D.length < 3) {
            return null;
        }

        const centroid = points2D.reduce((acc, point) => ({
            x: acc.x + point.x,
            y: acc.y + point.y
        }), { x: 0, y: 0 });
        centroid.x /= points2D.length;
        centroid.y /= points2D.length;

        const labelPositions = points2D.map((point) => {
            const dx = point.x - centroid.x;
            const dy = point.y - centroid.y;
            const length = Math.hypot(dx, dy) || 1;
            const offset = 52;
            return {
                x: point.x + (dx / length) * offset,
                y: point.y + (dy / length) * offset
            };
        });

        const sideEdgePairs = points2D.map((point, index) => [point, points2D[(index + 1) % points2D.length]]);
        const sideLabelPositions = sideEdgePairs.map(([a, b]) => this.getOffsetMidpoint(a, b, centroid, 34));
        const sideLabelAngles = sideEdgePairs.map(([a, b]) => {
            let angle = Math.atan2(b.y - a.y, b.x - a.x) * 180 / Math.PI;
            if (angle > 90) angle -= 180;
            if (angle < -90) angle += 180;
            return angle;
        });

        return {
            points2D,
            labelPositions,
            sideLabelPositions,
            sideLabelAngles,
            rightAnglePath: this.buildPolygonRightAnglePath(points2D),
            stageWidth,
            stageHeight
        };
    }

    buildTransformedExtractionLayout(baseLayout, quarterTurns = 0, isFlipped = false) {
        if (!baseLayout?.points2D?.length) {
            return null;
        }

        // Use the base layout's own stage dimensions so rotate/flip operations never
        // re-read the live DOM (which may have just been updated), preventing the
        // feedback loop that causes the modal height to creep up on each rotation.
        const stageWidth = baseLayout.stageWidth || 1000;
        const stageHeight = baseLayout.stageHeight || 760;

        const centroid = baseLayout.points2D.reduce((acc, point) => ({
            x: acc.x + point.x,
            y: acc.y + point.y
        }), { x: 0, y: 0 });
        centroid.x /= baseLayout.points2D.length;
        centroid.y /= baseLayout.points2D.length;

        const normalizedQuarterTurns = ((quarterTurns % 4) + 4) % 4;
        const transformed = baseLayout.points2D.map((point) => {
            let x = point.x - centroid.x;
            let y = point.y - centroid.y;

            if (isFlipped) {
                x *= -1;
            }

            for (let index = 0; index < normalizedQuarterTurns; index += 1) {
                const nextX = -y;
                const nextY = x;
                x = nextX;
                y = nextY;
            }

            return { x, y };
        });

        const fittedPoints = this.fitExtractionPointsToStage(transformed, stageWidth, stageHeight);
        return fittedPoints
            ? this.buildExtractionLayoutFromPoints(fittedPoints, stageWidth, stageHeight)
            : null;
    }

    rotateTriangleExtractionLayout() {
        if (this.triangleExtractTransitionState !== 'open' || !this.activeTriangleExtraction?.baseLayout) {
            return;
        }

        const nextQuarterTurns = (((this.activeTriangleExtraction.orientationQuarterTurns || 0) + 1) % 4 + 4) % 4;
        this.applyTriangleExtractionOrientation(nextQuarterTurns, this.activeTriangleExtraction.orientationFlipped === true);
    }

    flipTriangleExtractionLayout() {
        if (this.triangleExtractTransitionState !== 'open' || !this.activeTriangleExtraction?.baseLayout) {
            return;
        }

        this.applyTriangleExtractionOrientation(
            this.activeTriangleExtraction.orientationQuarterTurns || 0,
            !(this.activeTriangleExtraction.orientationFlipped === true)
        );
    }

    applyTriangleExtractionOrientation(quarterTurns, isFlipped) {
        if (!this.activeTriangleExtraction?.baseLayout) {
            return;
        }

        const layout = this.buildTransformedExtractionLayout(this.activeTriangleExtraction.baseLayout, quarterTurns, isFlipped);
        if (!layout) {
            return;
        }

        this.activeTriangleExtraction.orientationQuarterTurns = ((quarterTurns % 4) + 4) % 4;
        this.activeTriangleExtraction.orientationFlipped = isFlipped === true;
        this.activeTriangleExtraction.layout = layout;
        this.populateTriangleExtractionModal(
            layout,
            this.activeTriangleExtraction.labels,
            this.activeTriangleExtraction.item,
            this.activeTriangleExtraction.color
        );
        this.updateTriangleExtractionOrientationButtons();
    }

    updateTriangleExtractionOrientationButtons() {
        const hasExtraction = !!this.activeTriangleExtraction && this.triangleExtractTransitionState === 'open';
        if (this.triangleExtractRotateBtn) {
            this.triangleExtractRotateBtn.disabled = !hasExtraction;
        }
        if (this.triangleExtractFlipBtn) {
            this.triangleExtractFlipBtn.disabled = !hasExtraction;
            this.triangleExtractFlipBtn.setAttribute('aria-pressed', this.activeTriangleExtraction?.orientationFlipped === true ? 'true' : 'false');
        }
    }

    buildPolygonExtractionLayout(rawPoints, targetAspectRatio = 1000 / 760) {
        if (!Array.isArray(rawPoints) || rawPoints.length < 3) {
            return null;
        }

        const stageWidth = 1000;
        const safeAspect = Number.isFinite(targetAspectRatio) && targetAspectRatio > 0.1
            ? targetAspectRatio
            : (1000 / 760);
        const stageHeight = Math.max(1, Math.round(stageWidth / safeAspect));
        const padding = this.getExtractionLayoutPadding(stageWidth, stageHeight);
        const paddingX = padding.x;
        const paddingY = padding.y;

        const candidates = this.getPolygonOrientationCandidates(rawPoints);
        if (candidates.length === 0) {
            return null;
        }

        let bestCandidate = null;
        candidates.forEach((candidate) => {
            const minX = Math.min(...candidate.points.map((point) => point.x));
            const maxX = Math.max(...candidate.points.map((point) => point.x));
            const minY = Math.min(...candidate.points.map((point) => point.y));
            const maxY = Math.max(...candidate.points.map((point) => point.y));
            const width = Math.max(maxX - minX, 1e-6);
            const height = Math.max(maxY - minY, 1e-6);
            const scale = Math.min((stageWidth - paddingX * 2) / width, (stageHeight - paddingY * 2) / height);

            if (!bestCandidate || scale > bestCandidate.scale + 1e-6) {
                bestCandidate = {
                    ...candidate,
                    minX,
                    minY,
                    maxY,
                    width,
                    height,
                    scale
                };
                return;
            }

            if (bestCandidate && Math.abs(scale - bestCandidate.scale) <= 1e-6 && candidate.baseLength > bestCandidate.baseLength) {
                bestCandidate = {
                    ...candidate,
                    minX,
                    minY,
                    maxY,
                    width,
                    height,
                    scale
                };
            }
        });

        if (!bestCandidate) {
            return null;
        }

        const points2D = this.fitExtractionPointsToStage(bestCandidate.points, stageWidth, stageHeight, paddingX, paddingY);
        return points2D
            ? this.buildExtractionLayoutFromPoints(points2D, stageWidth, stageHeight)
            : null;
    }

    buildTriangleExtractionLayout(worldPoints, targetAspectRatio = 1000 / 760) {
        const sideAB = worldPoints[0].distanceTo(worldPoints[1]);
        const sideBC = worldPoints[1].distanceTo(worldPoints[2]);
        const sideCA = worldPoints[2].distanceTo(worldPoints[0]);
        const baseLength = sideAB;

        if (baseLength <= 1e-6) {
            return null;
        }

        const cX = (sideCA * sideCA - sideBC * sideBC + baseLength * baseLength) / (2 * baseLength);
        const cY = Math.sqrt(Math.max(0, sideCA * sideCA - cX * cX));
        if (cY <= 1e-6) {
            return null;
        }

        const rawPoints = [
            { x: 0, y: 0 },
            { x: baseLength, y: 0 },
            { x: cX, y: cY }
        ];

        const sideSquares = [sideAB * sideAB, sideBC * sideBC, sideCA * sideCA].sort((a, b) => a - b);
        const rightAngleTolerance = Math.max(1e-6, sideSquares[2] * 0.0025);
        const isRightAngled = Math.abs(sideSquares[0] + sideSquares[1] - sideSquares[2]) <= rightAngleTolerance;

        if (isRightAngled) {
            const stageWidth = 1000;
            const safeAspect = Number.isFinite(targetAspectRatio) && targetAspectRatio > 0.1
                ? targetAspectRatio
                : (1000 / 760);
            const stageHeight = Math.max(1, Math.round(stageWidth / safeAspect));
            const padding = this.getExtractionLayoutPadding(stageWidth, stageHeight);
            const hypotenuseLength = Math.max(sideAB, sideBC, sideCA);
            const lengthTolerance = Math.max(1e-6, hypotenuseLength * 0.0035);

            const legBaseCandidates = this.getTriangleOrientationCandidates(rawPoints)
                .filter((candidate) => Math.abs(candidate.baseLength - hypotenuseLength) > lengthTolerance);

            let bestLayout = null;
            let bestScaleScore = -Infinity;
            legBaseCandidates.forEach((candidate) => {
                const fittedPoints = this.fitExtractionPointsToStage(candidate.points, stageWidth, stageHeight, padding.x, padding.y);
                if (!fittedPoints) {
                    return;
                }

                const minX = Math.min(...candidate.points.map((point) => point.x));
                const maxX = Math.max(...candidate.points.map((point) => point.x));
                const minY = Math.min(...candidate.points.map((point) => point.y));
                const maxY = Math.max(...candidate.points.map((point) => point.y));
                const width = Math.max(maxX - minX, 1e-6);
                const height = Math.max(maxY - minY, 1e-6);
                const scaleScore = Math.min((stageWidth - padding.x * 2) / width, (stageHeight - padding.y * 2) / height);

                if (scaleScore > bestScaleScore + 1e-6) {
                    bestScaleScore = scaleScore;
                    bestLayout = this.buildExtractionLayoutFromPoints(fittedPoints, stageWidth, stageHeight);
                }
            });

            if (bestLayout) {
                return bestLayout;
            }
        }

        return this.buildPolygonExtractionLayout(rawPoints, targetAspectRatio);
    }

    buildPlaneExtractionLayout(worldPoints, targetAspectRatio = 1000 / 760) {
        if (!Array.isArray(worldPoints) || worldPoints.length !== 4) {
            return null;
        }

        const origin = worldPoints[0];
        const uVector = worldPoints[1].clone().sub(origin);
        if (uVector.lengthSq() <= 1e-12) {
            return null;
        }

        let normal = null;
        for (let index = 2; index < worldPoints.length; index += 1) {
            const candidate = new THREE.Vector3().crossVectors(uVector, worldPoints[index].clone().sub(origin));
            if (candidate.lengthSq() > 1e-12) {
                normal = candidate.normalize();
                break;
            }
        }

        if (!normal) {
            return null;
        }

        const u = uVector.clone().normalize();
        const v = new THREE.Vector3().crossVectors(normal, u).normalize();
        const rawPoints = worldPoints.map((point) => {
            const relative = point.clone().sub(origin);
            return {
                x: relative.dot(u),
                y: relative.dot(v)
            };
        });

        return this.buildPolygonExtractionLayout(rawPoints, targetAspectRatio);
    }

    getOffsetMidpoint(a, b, centroid, offsetDistance) {
        const midpoint = {
            x: (a.x + b.x) / 2,
            y: (a.y + b.y) / 2
        };
        const dx = midpoint.x - centroid.x;
        const dy = midpoint.y - centroid.y;
        const length = Math.hypot(dx, dy) || 1;
        return {
            x: midpoint.x + (dx / length) * offsetDistance,
            y: midpoint.y + (dy / length) * offsetDistance
        };
    }

    buildPolygonRightAnglePath(points2D, tolerance = 0.03) {
        if (!Array.isArray(points2D) || points2D.length < 3) {
            return '';
        }

        const pathParts = [];
        for (let index = 0; index < points2D.length; index += 1) {
            const vertex = points2D[index];
            const previous = points2D[(index - 1 + points2D.length) % points2D.length];
            const next = points2D[(index + 1) % points2D.length];
            const v1x = previous.x - vertex.x;
            const v1y = previous.y - vertex.y;
            const v2x = next.x - vertex.x;
            const v2y = next.y - vertex.y;
            const mag1 = Math.hypot(v1x, v1y);
            const mag2 = Math.hypot(v2x, v2y);
            if (mag1 <= 1e-6 || mag2 <= 1e-6) {
                continue;
            }

            const cosTheta = (v1x * v2x + v1y * v2y) / (mag1 * mag2);
            if (Math.abs(cosTheta) > tolerance) {
                continue;
            }

            const armLength = Math.min(34, Math.min(mag1, mag2) * 0.25);
            const prevUnit = { x: v1x / mag1, y: v1y / mag1 };
            const nextUnit = { x: v2x / mag2, y: v2y / mag2 };
            const p1 = { x: vertex.x + prevUnit.x * armLength, y: vertex.y + prevUnit.y * armLength };
            const p2 = { x: p1.x + nextUnit.x * armLength, y: p1.y + nextUnit.y * armLength };
            const p3 = { x: vertex.x + nextUnit.x * armLength, y: vertex.y + nextUnit.y * armLength };
            pathParts.push(`M ${p1.x} ${p1.y} L ${p2.x} ${p2.y} L ${p3.x} ${p3.y}`);
        }

        return pathParts.join(' ');
    }

    populateTriangleExtractionModal(layout, labels, item, color) {
        const pointIds = item.definition.pointIds;
        const pointSequence = this.formatPointSequence(pointIds);
        const pointCount = pointIds.length;
        const shapeName = item.type === 'plane' ? 'Quadrilateral' : 'Triangle';
        const polygonPoints = layout.points2D.map((point) => `${point.x},${point.y}`).join(' ');
        const pointKeys = ['A', 'B', 'C', 'D'];
        const activePointKeys = pointKeys.slice(0, pointCount);
        const sideKeys = activePointKeys.map((key, index) => `${key}${activePointKeys[(index + 1) % activePointKeys.length]}`);

        if (this.triangleExtractTitle) {
            this.triangleExtractTitle.textContent = `${shapeName} ${pointSequence}`;
        }
        if (this.triangleExtractCloseBtn) {
            this.triangleExtractCloseBtn.setAttribute('aria-label', `Close inspected ${shapeName.toLowerCase()}`);
        }
        if (this.triangleExtractSvg) {
            this.triangleExtractSvg.setAttribute('viewBox', `0 0 ${layout.stageWidth} ${layout.stageHeight}`);
            this.triangleExtractSvg.setAttribute('aria-label', `${shapeName} 2D view`);
        }
        if (this.triangleExtractPolygon) {
            this.triangleExtractPolygon.setAttribute('points', polygonPoints);
            if (color) this.triangleExtractPolygon.style.fill = color.fill;
        }
        if (this.triangleExtractOutline) {
            this.triangleExtractOutline.setAttribute('points', `${polygonPoints} ${layout.points2D[0].x},${layout.points2D[0].y}`);
            if (color) this.triangleExtractOutline.style.stroke = color.stroke;
        }
        if (this.triangleExtractRightAngle) {
            const hasRightAngle = layout.rightAnglePath.length > 0;
            this.triangleExtractRightAngle.hidden = !hasRightAngle;
            this.triangleExtractRightAngle.setAttribute('d', layout.rightAnglePath);
            if (color) this.triangleExtractRightAngle.style.stroke = color.stroke;
        }

        pointKeys.forEach((key, index) => {
            const labelEl = this.triangleExtractLabelEls[key];
            if (!labelEl) return;
            if (index >= pointCount) {
                labelEl.textContent = '';
                labelEl.setAttribute('visibility', 'hidden');
                return;
            }
            labelEl.textContent = labels[index];
            labelEl.setAttribute('x', `${layout.labelPositions[index].x}`);
            labelEl.setAttribute('y', `${layout.labelPositions[index].y}`);
            labelEl.removeAttribute('visibility');
        });

        Object.entries(this.triangleExtractSideEls).forEach(([key, sideEl]) => {
            if (!sideEl) return;
            sideEl.textContent = '';
            sideEl.setAttribute('visibility', 'hidden');
            sideEl.removeAttribute('transform');
        });

        const sidePairs = this.getPolygonEdgePairsInOrder(pointIds);
        sideKeys.forEach((key, index) => {
            const sideEl = this.triangleExtractSideEls[key];
            if (!sideEl) return;
            const edgeLabelObj = this.findEdgeLabelObject(sidePairs[index]);
            const labelText = edgeLabelObj?.definition?.text ?? '';
            if (labelText) {
                const pos = layout.sideLabelPositions[index];
                const angle = layout.sideLabelAngles[index];
                sideEl.textContent = labelText;
                sideEl.setAttribute('x', `${pos.x}`);
                sideEl.setAttribute('y', `${pos.y}`);
                sideEl.setAttribute('transform', `rotate(${angle}, ${pos.x}, ${pos.y})`);
                sideEl.removeAttribute('visibility');
            } else {
                sideEl.textContent = '';
                sideEl.setAttribute('visibility', 'hidden');
                sideEl.removeAttribute('transform');
            }
        });

        if (this.triangleExtractAnglesGroup) {
            this.triangleExtractAnglesGroup.innerHTML = '';
            const ns = 'http://www.w3.org/2000/svg';
            const angleObjects = this.findAngleLabelObjectsForPolygon(pointIds);
            for (const angleObj of angleObjects) {
                const def = angleObj.definition;
                const vertexIndex = pointIds.indexOf(def.pointIds[1]);
                const aIndex = pointIds.indexOf(def.pointIds[0]);
                const cIndex = pointIds.indexOf(def.pointIds[2]);
                if (vertexIndex < 0 || aIndex < 0 || cIndex < 0) continue;

                const vertex = layout.points2D[vertexIndex];
                const aPoint = layout.points2D[aIndex];
                const cPoint = layout.points2D[cIndex];
                const d1x = aPoint.x - vertex.x, d1y = aPoint.y - vertex.y;
                const d2x = cPoint.x - vertex.x, d2y = cPoint.y - vertex.y;
                const armLen1 = Math.hypot(d1x, d1y);
                const armLen2 = Math.hypot(d2x, d2y);
                if (armLen1 < 1e-6 || armLen2 < 1e-6) continue;

                const dir1 = { x: d1x / armLen1, y: d1y / armLen1 };
                const dir2 = { x: d2x / armLen2, y: d2y / armLen2 };
                const isCompactViewport = window.matchMedia?.('(max-width: 768px)').matches ?? (window.innerWidth < 768);
                const arcRadiusFactor = isCompactViewport ? 0.34 : 0.25;
                const maxArcRadius = isCompactViewport ? 92 : 65;
                const r = Math.min(maxArcRadius, Math.min(armLen1, armLen2) * arcRadiusFactor);

                const startX = vertex.x + dir1.x * r;
                const startY = vertex.y + dir1.y * r;
                const endX = vertex.x + dir2.x * r;
                const endY = vertex.y + dir2.y * r;
                const cross = dir1.x * dir2.y - dir1.y * dir2.x;
                const sweep = cross > 0 ? 1 : 0;

                const arcColor = this.getTriangleExtractionColor(Number.isFinite(def.color) ? def.color : 0x00d1b2).stroke;

                const arcEl = document.createElementNS(ns, 'path');
                arcEl.setAttribute('d', `M ${startX} ${startY} A ${r} ${r} 0 0 ${sweep} ${endX} ${endY}`);
                arcEl.setAttribute('class', 'triangle-extract-angle-arc');
                arcEl.style.stroke = arcColor;
                this.triangleExtractAnglesGroup.appendChild(arcEl);

                const bisX = dir1.x + dir2.x;
                const bisY = dir1.y + dir2.y;
                const bisMag = Math.hypot(bisX, bisY) || 1;
                const labelDist = r + (isCompactViewport ? 36 : 28);
                const textEl = document.createElementNS(ns, 'text');
                textEl.setAttribute('x', `${vertex.x + (bisX / bisMag) * labelDist}`);
                textEl.setAttribute('y', `${vertex.y + (bisY / bisMag) * labelDist}`);
                textEl.setAttribute('class', 'triangle-extract-angle-text');
                textEl.textContent = def.text;
                this.triangleExtractAnglesGroup.appendChild(textEl);
            }
        }
    }

    formatTriangleSideLength(lengthValue) {
        const rounded = Math.round(lengthValue * 100) / 100;
        return Number.isInteger(rounded) ? `${rounded}` : rounded.toFixed(2).replace(/0$/, '').replace(/\.$/, '');
    }

    parsePositiveInteger(value) {
        const text = `${value ?? ''}`.trim();
        if (!/^[1-9]\d*$/.test(text)) {
            return null;
        }

        const parsed = Number(text);
        return Number.isSafeInteger(parsed) ? parsed : null;
    }

    greatestCommonDivisor(leftValue, rightValue) {
        let left = Math.abs(leftValue);
        let right = Math.abs(rightValue);

        while (right !== 0) {
            const remainder = left % right;
            left = right;
            right = remainder;
        }

        return left || 1;
    }

    reduceRatio(leftValue, rightValue) {
        const left = this.parsePositiveInteger(leftValue);
        const right = this.parsePositiveInteger(rightValue);
        if (!left || !right) {
            return null;
        }

        const divisor = this.greatestCommonDivisor(left, right);
        return {
            left: left / divisor,
            right: right / divisor
        };
    }

    toggleGrid() {
        this.gridVisible = !this.gridVisible;
        if (this.grid) {
            this.grid.visible = this.gridVisible;
        }
        this.updateGridToggleUI();
    }

    updateGridToggleUI() {
        if (!this.gridToggleBtn) {
            return;
        }

        this.gridToggleBtn.style.background = this.gridVisible ? '#2A3F5A' : '#1a2a3f';
        this.gridToggleBtn.style.opacity = this.gridVisible ? '1' : '0.6';
        this.gridToggleBtn.title = this.gridVisible
            ? 'Grid enabled (click to disable)'
            : 'Grid disabled (click to enable)';

        if (this.gridIcon) {
            this.gridIcon.classList.toggle('grid-active', this.gridVisible);
        }
    }

    togglePointMarkers() {
        this.pointMarkersVisible = !this.pointMarkersVisible;
        this.updatePointMarkerToggleUI();
        this.pointMarkers.forEach((marker) => { marker.visible = this.pointMarkersVisible; });
    }

    updatePointMarkerToggleUI() {
        if (!this.pointMarkerToggleBtn) return;
        this.pointMarkerToggleBtn.title = this.pointMarkersVisible
            ? 'Point markers visible (click to hide)'
            : 'Point markers hidden (click to show)';
        if (this.pointFilledIcon) {
            this.pointFilledIcon.classList.toggle('point-marker-active', this.pointMarkersVisible);
        }
        if (this.pointHollowIcon) {
            this.pointHollowIcon.classList.toggle('point-marker-active', !this.pointMarkersVisible);
        }
    }

    updateGhostToggleUI() {
        if (!this.ghostToggleBtn) {
            return;
        }

        this.ghostToggleBtn.title = this.ghostFaces
            ? 'Ghost faces enabled (click to disable)'
            : 'Solid faces enabled (click to enable ghost mode)';

        if (this.ghostWireIcon) {
            this.ghostWireIcon.classList.toggle('ghost-active', this.ghostFaces);
        }

        if (this.ghostSolidIcon) {
            this.ghostSolidIcon.classList.toggle('ghost-active', !this.ghostFaces);
        }
    }

    toggleLabelBadgeMode() {
        const cycle = { badge: 'off', off: 'plain', plain: 'badge' };
        this.labelMode = cycle[this.labelMode] || 'badge';
        this.updateLabelBadgeToggleUI();
        this.refreshSceneTextSizing();
    }

    updateLabelBadgeToggleUI() {
        if (!this.labelBadgeToggleBtn) {
            return;
        }

        const titles = {
            badge: 'Labels with badges (click to hide)',
            off: 'Labels hidden (click for plain)',
            plain: 'Plain labels (click to restore badges)'
        };
        this.labelBadgeToggleBtn.title = titles[this.labelMode] || '';

        if (this.labelPlainIcon) {
            this.labelPlainIcon.classList.toggle('label-badge-active', this.labelMode === 'plain');
        }
        if (this.labelBadgeIcon) {
            this.labelBadgeIcon.classList.toggle('label-badge-active', this.labelMode === 'badge');
        }
        if (this.labelOffIcon) {
            this.labelOffIcon.classList.toggle('label-badge-active', this.labelMode === 'off');
        }
    }

    toggleDisplaySizeMode() {
        this.displaySizeMode = this.displaySizeMode === 'small' ? 'large' : 'small';
        this.updateDisplaySizeToggleUI();
        this.refreshSceneTextSizing();
    }

    updateDisplaySizeToggleUI() {
        if (this.sizeSmallOption) {
            this.sizeSmallOption.classList.toggle('size-active', this.displaySizeMode === 'small');
        }
        if (this.sizeLargeOption) {
            this.sizeLargeOption.classList.toggle('size-active', this.displaySizeMode === 'large');
        }
    }

    refreshSceneTextSizing() {
        this.rebuildConstructions();
        this.buildPointMarkers();
        this.updatePointMarkerStyles();
        this.syncAllLabelVisibility();
    }

    getEdgeColor() {
        return this.themeMode === 'dark' ? 0xffffff : 0x000000;
    }

    getLabelTextColor() {
        return (this.labelMode !== 'badge' && this.themeMode === 'dark') ? '#ffffff' : '#000000';
    }

    toggleThemeMode() {
        this.themeMode = this.themeMode === 'light' ? 'dark' : 'light';
        this.applyThemeMode();
    }

    applyThemeMode() {
        document.documentElement.setAttribute('data-theme', this.themeMode);
        if (this.scene?.background) {
            this.scene.background.set(this.themeMode === 'dark' ? 0x606060 : 0xffffff);
        }
        this.updateGridThemeAppearance();
        this.updateThemeToggleUI();
        const edgeColor = this.getEdgeColor();
        this.primitiveGroup?.traverse((obj) => {
            if (obj.material instanceof THREE.LineBasicMaterial) {
                obj.material.color.setHex(edgeColor);
                obj.material.needsUpdate = true;
            }
        });
        this.buildPointMarkers();
    }

    updateThemeToggleUI() {
        if (this.lightIcon) {
            this.lightIcon.classList.toggle('theme-active', this.themeMode === 'light');
        }
        if (this.darkIcon) {
            this.darkIcon.classList.toggle('theme-active', this.themeMode === 'dark');
        }
    }

    updateGridThemeAppearance() {
        if (!this.grid) {
            return;
        }

        const materials = Array.isArray(this.grid.material)
            ? this.grid.material
            : [this.grid.material];
        const isDark = this.themeMode === 'dark';

        materials.forEach((material) => {
            if (!material) {
                return;
            }

            material.transparent = isDark;
            material.opacity = isDark ? 0.35 : 1;
            material.needsUpdate = true;
        });
    }

    onWindowResize() {
        const width = this.canvas.clientWidth;
        const height = this.canvas.clientHeight;
        this.camera.aspect = width / height;
        this.camera.updateProjectionMatrix();
        this.renderer.setSize(width, height, false);
        this.updateConstructionLineMaterialResolutions();
        this.refreshSliderBadges();
        if (this.isTriangleExtractionOpen() && this.triangleExtractModal) {
            this.triangleExtractModal.style.transform = 'translate(0px, 0px) scale(1)';
        }
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

        const existingPrimitives = new Set(this.compositeSlots.map((s) => s.primitive));
        const disallowedSelfAdds = new Set(['cuboid', 'cylinder']);

        return Object.keys(this.defaultParams).filter((primKey) => {
            if (existingPrimitives.has(primKey) && disallowedSelfAdds.has(primKey)) return false;
            const guestFaces = ATTACHMENT_FACES[primKey] || [];
            if (guestFaces.length === 0) return false;

            const draftSlot = {
                primitive: primKey,
                orientation: this.orientations[primKey]?.[0]?.value || 'standard',
                params: { ...this.defaultParams[primKey] }
            };

            return this.getValidHostFaceEntries(draftSlot, this.compositeSlots, { excludeOccupied: true }).length > 0;
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
            const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal, hostFaceDef.type);
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
        const picked = availableEntries[0] || null;
        if (!picked) {
            slot.hostSlotId = null;
            slot.hostFaceId = null;
            return null;
        }

        slot.hostSlotId = picked.slotId;
        slot.hostFaceId = picked.faceId;
        return picked;
    }

    getGuestAttachFaceDef(guestSlot, hostFaceNormal, hostFaceType = null) {
        const faces = ATTACHMENT_FACES[guestSlot.primitive] || [];
        if (faces.length === 0) return null;

        if (guestSlot?.attachFaceId) {
            const explicitFace = faces.find((face) => face.id === guestSlot.attachFaceId);
            if (explicitFace && (!hostFaceType || explicitFace.type === hostFaceType)) {
                return explicitFace;
            }
        }

        const targetNormal = hostFaceNormal.clone().negate();
        const ABS_DOT_EPSILON = 1e-6;
        const DOT_EPSILON = 1e-6;

        const pickBest = (candidates) => {
            let best = null;
            let bestIndex = Infinity;
            let bestAbsDot = -Infinity;
            let bestDot = -Infinity;

            candidates.forEach(({ face, index }) => {
                const faceNormal = this.resolveFaceNormal(face, guestSlot.params);
                const dot = faceNormal.dot(targetNormal);
                const absDot = Math.abs(dot);

                if (best === null) {
                    best = face; bestIndex = index; bestAbsDot = absDot; bestDot = dot;
                    return;
                }

                const absDelta = absDot - bestAbsDot;
                if (absDelta > ABS_DOT_EPSILON) {
                    best = face; bestIndex = index; bestAbsDot = absDot; bestDot = dot;
                    return;
                }

                if (Math.abs(absDelta) <= ABS_DOT_EPSILON) {
                    const dotDelta = dot - bestDot;
                    if (dotDelta > DOT_EPSILON || (Math.abs(dotDelta) <= DOT_EPSILON && index < bestIndex)) {
                        best = face; bestIndex = index; bestAbsDot = absDot; bestDot = dot;
                    }
                }
            });
            return best;
        };

        const indexed = faces.map((face, index) => ({ face, index }));

        // Prefer matching face type so (e.g.) a rectangle host never picks a triangle guest face
        if (hostFaceType) {
            const sameType = indexed.filter(({ face }) => face.type === hostFaceType);
            if (sameType.length > 0) {
                return pickBest(sameType);
            }
        }

        return pickBest(indexed);
    }

    getAttachmentQuarterTurns(slot, guestFaceDef = null, hostFaceDef = null) {
        if (!slot || slot.primitive !== 'right-triangle-prism') {
            return 0;
        }

        if (guestFaceDef?.type !== 'rectangle' || hostFaceDef?.type !== 'rectangle') {
            return 0;
        }

        return ((slot.attachRotationQuarterTurns || 0) % 4 + 4) % 4;
    }

    getAttachmentDimsForHost(slot, guestFaceDef, hostFaceDef = null) {
        const guestDims = this.resolveFaceDims(guestFaceDef, slot.params).slice();
        const turns = this.getAttachmentQuarterTurns(slot, guestFaceDef, hostFaceDef);
        if (guestDims.length >= 2 && turns % 2 === 1) {
            [guestDims[0], guestDims[1]] = [guestDims[1], guestDims[0]];
        }
        return guestDims;
    }

    snapSlotDimensions(guestSlot) {
        const guestSlotIndex = this.compositeSlots.findIndex((s) => s.id === guestSlot.id);
        if (guestSlotIndex <= 0) return;
        const prevSlots = this.compositeSlots.slice(0, guestSlotIndex);
        const entry = this.ensureSlotHostBinding(guestSlot, prevSlots);
        if (!entry) return;

        const { slot: hostSlot, faceDef: hostFaceDef } = entry;
    this.syncAttachmentSpecificVariants(guestSlot, hostSlot, hostFaceDef);
        const hostFaceNormal = this.resolveFaceNormal(hostFaceDef, hostSlot.params);
        const guestFaceDef = this.getGuestAttachFaceDef(guestSlot, hostFaceNormal, hostFaceDef.type);
        if (!guestFaceDef) return;

        const hostDims = this.resolveFaceDims(hostFaceDef, hostSlot.params);
        const guestDims = this.getAttachmentDimsForHost(guestSlot, guestFaceDef, hostFaceDef);
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

    syncAttachmentSpecificVariants(guestSlot, hostSlot, hostFaceDef) {
        if (!guestSlot || !hostSlot || !hostFaceDef) {
            return;
        }

        const isGuestTetra = guestSlot.primitive === 'tetrahedron';
        const isHostTetra = hostSlot.primitive === 'tetrahedron';
        const isGuestPrism = guestSlot.primitive === 'right-triangle-prism';
        const isHostPrism = hostSlot.primitive === 'right-triangle-prism';

        // Prism <-> tetrahedron attachments should always drive tetra mode from prism mode,
        // regardless of which primitive is host/guest.
        if (isGuestTetra && isHostPrism && (hostFaceDef.id === 'front-triangle' || hostFaceDef.id === 'back-triangle')) {
            const prismMode = normalizeTriangularPrismMode(hostSlot.params.triangleMode);
            guestSlot.params.baseTriangleMode = prismMode === 'isosceles' ? 'isosceles' : 'right-angled';
            guestSlot.params.baseMirror = (prismMode === 'right-above-A') !== (hostFaceDef.id === 'back-triangle');
            return;
        }

        if (isGuestPrism && isHostTetra && hostFaceDef.id === 'base-triangle') {
            const hostFaceNormal = this.resolveFaceNormal(hostFaceDef, hostSlot.params);
            const prismFaceDef = this.getGuestAttachFaceDef(guestSlot, hostFaceNormal, hostFaceDef.type);
            if (prismFaceDef && (prismFaceDef.id === 'front-triangle' || prismFaceDef.id === 'back-triangle')) {
                const prismMode = normalizeTriangularPrismMode(guestSlot.params.triangleMode);
                hostSlot.params.baseTriangleMode = prismMode === 'isosceles' ? 'isosceles' : 'right-angled';
                hostSlot.params.baseMirror = (prismMode === 'right-above-A') !== (prismFaceDef.id === 'back-triangle');
            }
            return;
        }

        // Tetrahedron <-> tetrahedron keeps mirrored right-angle compatibility.
        if (isGuestTetra && isHostTetra && hostFaceDef.id === 'base-triangle') {
            const hostBaseMode = normalizeTetrahedronBaseMode(hostSlot.params.baseTriangleMode);
            guestSlot.params.baseTriangleMode = hostBaseMode;
            guestSlot.params.baseMirror = hostBaseMode === 'right-angled';
            return;
        }

        if (isGuestTetra) {
            guestSlot.params.baseMirror = false;
        }
    }

    isTetraBaseModeLocked(slot) {
        return !!this.getLinkedTetraBaseController(slot);
    }

    getLinkedTetraBaseController(slot) {
        if (!slot || slot.primitive !== 'tetrahedron') {
            return null;
        }

        if (slot.hostSlotId != null && slot.hostFaceId) {
            const hostSlot = this.compositeSlots.find((s) => s.id === slot.hostSlotId);
            const hostFaceDef = this.getFaceDefById(hostSlot, slot.hostFaceId);
            if (hostSlot?.primitive === 'right-triangle-prism'
                && (hostFaceDef?.id === 'front-triangle' || hostFaceDef?.id === 'back-triangle')) {
                return hostSlot;
            }

            if (hostSlot?.primitive === 'tetrahedron' && hostFaceDef?.id === 'base-triangle') {
                return hostSlot;
            }
        }

        const attachedPrismGuest = this.compositeSlots.find((child) => {
            if (child.hostSlotId !== slot.id || child.primitive !== 'right-triangle-prism' || !child.hostFaceId) {
                return false;
            }

            const hostFaceDef = this.getFaceDefById(slot, child.hostFaceId);
            return hostFaceDef?.id === 'base-triangle';
        });

        return attachedPrismGuest || null;
    }

    syncLinkedTetraModesForPrism(prismSlot) {
        if (!prismSlot || prismSlot.primitive !== 'right-triangle-prism') {
            return;
        }

        // Prism as guest attached to a tetrahedron host.
        if (prismSlot.hostSlotId != null && prismSlot.hostFaceId) {
            const hostSlot = this.compositeSlots.find((s) => s.id === prismSlot.hostSlotId);
            const hostFaceDef = this.getFaceDefById(hostSlot, prismSlot.hostFaceId);
            if (hostSlot?.primitive === 'tetrahedron' && hostFaceDef?.id === 'base-triangle') {
                this.syncAttachmentSpecificVariants(prismSlot, hostSlot, hostFaceDef);
            }
        }

        // Prism as host with tetrahedron guests.
        this.compositeSlots.forEach((child) => {
            if (child.hostSlotId !== prismSlot.id || child.primitive !== 'tetrahedron' || !child.hostFaceId) {
                return;
            }

            const hostFaceDef = this.getFaceDefById(prismSlot, child.hostFaceId);
            if (hostFaceDef?.id === 'front-triangle' || hostFaceDef?.id === 'back-triangle') {
                this.syncAttachmentSpecificVariants(child, prismSlot, hostFaceDef);
            }
        });
    }

    normalizeCuboidPyramidHostOrder() {
        if (this.compositeSlots.length !== 2) {
            return;
        }

        const cuboidSlot = this.compositeSlots.find((slot) => slot.primitive === 'cuboid');
        const pyramidSlot = this.compositeSlots.find((slot) => slot.primitive === 'rectangular-pyramid');
        if (!cuboidSlot || !pyramidSlot) {
            return;
        }

        // Already in preferred directed attachment: cuboid host -> pyramid guest.
        if (cuboidSlot.hostSlotId == null && pyramidSlot.hostSlotId === cuboidSlot.id) {
            return;
        }

        this.compositeSlots = [cuboidSlot, pyramidSlot];
        cuboidSlot.hostSlotId = null;
        cuboidSlot.hostFaceId = null;
        pyramidSlot.hostSlotId = null;
        pyramidSlot.hostFaceId = null;
        this.snapSlotDimensions(pyramidSlot);
    }

    normalizeCuboidPrismHostOrder() {
        if (this.compositeSlots.length !== 2) {
            return;
        }

        const cuboidSlot = this.compositeSlots.find((slot) => slot.primitive === 'cuboid');
        const prismSlot = this.compositeSlots.find((slot) => slot.primitive === 'right-triangle-prism');
        if (!cuboidSlot || !prismSlot) {
            return;
        }

        // Already in preferred directed attachment: cuboid host -> prism guest.
        if (cuboidSlot.hostSlotId == null && prismSlot.hostSlotId === cuboidSlot.id) {
            return;
        }

        this.compositeSlots = [cuboidSlot, prismSlot];
        cuboidSlot.hostSlotId = null;
        cuboidSlot.hostFaceId = null;
        prismSlot.hostSlotId = null;
        prismSlot.hostFaceId = null;
        this.snapSlotDimensions(prismSlot);
    }

    normalizeCylinderConeHostOrder() {
        if (this.compositeSlots.length !== 2) {
            return;
        }

        const cylinderSlot = this.compositeSlots.find((slot) => slot.primitive === 'cylinder');
        const coneSlot = this.compositeSlots.find((slot) => slot.primitive === 'cone');
        if (!cylinderSlot || !coneSlot) {
            return;
        }

        // Already in preferred directed attachment: cylinder host -> cone guest.
        if (cylinderSlot.hostSlotId == null && coneSlot.hostSlotId === cylinderSlot.id) {
            return;
        }

        this.compositeSlots = [cylinderSlot, coneSlot];
        cylinderSlot.hostSlotId = null;
        cylinderSlot.hostFaceId = null;
        coneSlot.hostSlotId = null;
        coneSlot.hostFaceId = null;
        this.snapSlotDimensions(coneSlot);
    }

    normalizeCylinderHemisphereHostOrder() {
        if (this.compositeSlots.length !== 2) {
            return;
        }

        const cylinderSlot = this.compositeSlots.find((slot) => slot.primitive === 'cylinder');
        const hemisphereSlot = this.compositeSlots.find((slot) => slot.primitive === 'hemisphere');
        if (!cylinderSlot || !hemisphereSlot) {
            return;
        }

        // Already in preferred directed attachment: cylinder host -> hemisphere guest.
        if (cylinderSlot.hostSlotId == null && hemisphereSlot.hostSlotId === cylinderSlot.id) {
            return;
        }

        this.compositeSlots = [cylinderSlot, hemisphereSlot];
        cylinderSlot.hostSlotId = null;
        cylinderSlot.hostFaceId = null;
        hemisphereSlot.hostSlotId = null;
        hemisphereSlot.hostFaceId = null;
        this.snapSlotDimensions(hemisphereSlot);
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
                attachFaceId: null,
                attachRotationQuarterTurns: 0,
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
                attachFaceId: null,
                attachRotationQuarterTurns: 0,
            };

            const prevSlots = this.compositeSlots.slice();
            const availableEntries = this.getValidHostFaceEntries(slot, prevSlots, { excludeOccupied: true });
            if (availableEntries.length === 0) return;

            this.compositeSlots.push(slot);
            this.snapSlotDimensions(slot);
            this.normalizeCuboidPyramidHostOrder();
            this.normalizeCuboidPrismHostOrder();
            this.normalizeCylinderConeHostOrder();
            this.normalizeCylinderHemisphereHostOrder();
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

    getCurrentHostEntryForSlot(slot) {
        if (!slot || slot.hostSlotId == null || !slot.hostFaceId) {
            return null;
        }

        const hostSlot = this.compositeSlots.find((candidate) => candidate.id === slot.hostSlotId);
        const hostFaceDef = this.getFaceDefById(hostSlot, slot.hostFaceId);
        if (!hostSlot || !hostFaceDef) {
            return null;
        }

        return { slot: hostSlot, faceDef: hostFaceDef };
    }

    getPrismAttachFaceCandidates(slot, hostFaceType = 'rectangle') {
        if (!slot || slot.primitive !== 'right-triangle-prism') {
            return [];
        }

        return (ATTACHMENT_FACES[slot.primitive] || []).filter((faceDef) => faceDef.type === hostFaceType);
    }

    getPrismAttachFaceLabel(faceId) {
        return (ATTACHMENT_FACES['right-triangle-prism'] || []).find((faceDef) => faceDef.id === faceId)?.label || 'Rectangle';
    }

    isPrismRectAttachmentConfigurable(slot) {
        if (!slot || slot.primitive !== 'right-triangle-prism' || slot.hostSlotId == null || !slot.hostFaceId) {
            return false;
        }

        const hostEntry = this.getCurrentHostEntryForSlot(slot);
        return hostEntry?.faceDef?.type === 'rectangle';
    }

    cyclePrismAttachFace(slotId) {
        const slot = this.compositeSlots.find((candidate) => candidate.id === slotId);
        if (!this.isPrismRectAttachmentConfigurable(slot)) return;

        const hostEntry = this.getCurrentHostEntryForSlot(slot);
        const candidates = this.getPrismAttachFaceCandidates(slot, hostEntry.faceDef.type);
        if (candidates.length <= 1) return;

        const hostFaceNormal = this.resolveFaceNormal(hostEntry.faceDef, hostEntry.slot.params);
        const currentFace = this.getGuestAttachFaceDef(slot, hostFaceNormal, hostEntry.faceDef.type);
        const currentIndex = candidates.findIndex((candidate) => candidate.id === currentFace?.id);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % candidates.length : 0;

        slot.attachFaceId = candidates[nextIndex].id;
        this.snapSlotDimensions(slot);
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getPrismAttachRotationLabel(quarterTurns) {
        const normalized = ((quarterTurns || 0) % 4 + 4) % 4;
        return `${normalized * 90}\u00b0`;
    }

    cyclePrismAttachRotation(slotId) {
        const slot = this.compositeSlots.find((candidate) => candidate.id === slotId);
        if (!this.isPrismRectAttachmentConfigurable(slot)) return;

        slot.attachRotationQuarterTurns = (((slot.attachRotationQuarterTurns || 0) + 1) % 4 + 4) % 4;
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
        this.syncLinkedTetraModesForPrism(slot);

        const slotIndex = this.compositeSlots.findIndex((s) => s.id === slotId);
        this.compositeSlots.slice(slotIndex + 1).forEach((laterSlot) => this.snapSlotDimensions(laterSlot));

        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    toggleCuboidFaceCenters(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'cuboid') return;

        const modes = ['off', 'top-bottom', 'left-right', 'front-back', 'all'];
        const current = this.getCuboidFaceCentersMode(slot.params);
        const currentIndex = modes.indexOf(current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % modes.length : 0;
        slot.params.includeFaceCentersMode = modes[nextIndex];
        this.resetSceneObjects();
        this.buildComposite();
        this.renderCompositeCards();
    }

    getCuboidFaceCentersMode(params = {}) {
        const mode = params.includeFaceCentersMode;
        if (mode === 'off' || mode === 'top-bottom' || mode === 'left-right' || mode === 'front-back' || mode === 'all') {
            return mode;
        }

        // Backward compatibility with legacy boolean flag.
        return params.includeFaceCenters === true ? 'all' : 'off';
    }

    getCuboidFaceCentersModeLabel(mode) {
        const labels = {
            off: 'Off',
            'top-bottom': 'Top/Bottom',
            'left-right': 'Left/Right',
            'front-back': 'Front/Back',
            all: 'All'
        };
        return labels[mode] || 'Off';
    }

    getTetrahedronTriangleModeLabel(mode) {
        return this.tetrahedronTriangleModes.find((opt) => opt.value === normalizeTetrahedronBaseMode(mode))?.label || 'Isosceles';
    }

    cycleTetrahedronTriangleMode(slotId) {
        const slot = this.compositeSlots.find((s) => s.id === slotId);
        if (!slot || slot.primitive !== 'tetrahedron') return;
        const linkedController = this.getLinkedTetraBaseController(slot);
        if (linkedController?.primitive === 'right-triangle-prism') {
            this.cycleTriangularPrismMode(linkedController.id);
            return;
        }

        if (linkedController?.primitive === 'tetrahedron') {
            this.cycleTetrahedronTriangleMode(linkedController.id);
            return;
        }

        const options = this.tetrahedronTriangleModes;
        const current = normalizeTetrahedronBaseMode(slot.params.baseTriangleMode);
        const currentIndex = options.findIndex((opt) => opt.value === current);
        const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % options.length : 0;
        slot.params.baseTriangleMode = options[nextIndex].value;
        this.snapSlotDimensions(slot);

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

    isOrientationControlVisible(slot) {
        if (!slot) {
            return false;
        }

        if (slot.hostSlotId != null) {
            const hostSlot = this.compositeSlots.find((candidate) => candidate.id === slot.hostSlotId);
            const hostFaceDef = this.getFaceDefById(hostSlot, slot.hostFaceId);

            if (slot.primitive === 'rectangular-pyramid' && hostSlot?.primitive === 'cuboid') {
                // In this attachment context apex-up/down is visually equivalent, so hide inert controls.
                return false;
            }

            if ((slot.primitive === 'cone' || slot.primitive === 'cylinder') && hostFaceDef?.type === 'circle') {
                // Circular joins are rotationally symmetric; these controls are redundant when attached.
                return false;
            }
        }

        return true;
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
        this.updatePrimarySectionCounts();

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
            if (orientOptions.length > 1 && this.isOrientationControlVisible(slot)) {
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

            if (slot.primitive === 'cuboid') {
                const cuboidFaceCentersBtn = document.createElement('button');
                cuboidFaceCentersBtn.type = 'button';
                cuboidFaceCentersBtn.className = 'card-cycle-btn';
                cuboidFaceCentersBtn.dataset.toggleCuboidFaceCentersSlotId = String(slot.id);
                const centersMode = this.getCuboidFaceCentersMode(slot.params);
                cuboidFaceCentersBtn.textContent = `Face Centres: ${this.getCuboidFaceCentersModeLabel(centersMode)}`;
                card.appendChild(cuboidFaceCentersBtn);
            }

            if (slot.primitive === 'right-triangle-prism') {
                const triangleModeCycleBtn = document.createElement('button');
                triangleModeCycleBtn.type = 'button';
                triangleModeCycleBtn.className = 'card-cycle-btn';
                triangleModeCycleBtn.dataset.cycleTriangleModeSlotId = String(slot.id);
                triangleModeCycleBtn.textContent = `Cycle: ${this.getTriangularPrismModeLabel(slot.params.triangleMode)}`;
                card.appendChild(triangleModeCycleBtn);

                if (this.isPrismRectAttachmentConfigurable(slot)) {
                    const hostEntry = this.getCurrentHostEntryForSlot(slot);
                    const hostFaceNormal = this.resolveFaceNormal(hostEntry.faceDef, hostEntry.slot.params);
                    const currentFace = this.getGuestAttachFaceDef(slot, hostFaceNormal, hostEntry.faceDef.type);

                    const attachFaceCycleBtn = document.createElement('button');
                    attachFaceCycleBtn.type = 'button';
                    attachFaceCycleBtn.className = 'card-cycle-btn';
                    attachFaceCycleBtn.dataset.cyclePrismAttachFaceSlotId = String(slot.id);
                    attachFaceCycleBtn.textContent = `Flush Face: ${this.getPrismAttachFaceLabel(currentFace?.id)}`;
                    card.appendChild(attachFaceCycleBtn);

                    const attachRotationCycleBtn = document.createElement('button');
                    attachRotationCycleBtn.type = 'button';
                    attachRotationCycleBtn.className = 'card-cycle-btn';
                    attachRotationCycleBtn.dataset.cyclePrismAttachRotationSlotId = String(slot.id);
                    attachRotationCycleBtn.textContent = `Rotate: ${this.getPrismAttachRotationLabel(slot.attachRotationQuarterTurns)}`;
                    card.appendChild(attachRotationCycleBtn);
                }
            }

            if (slot.primitive === 'tetrahedron') {
                const tetrahedronModeCycleBtn = document.createElement('button');
                tetrahedronModeCycleBtn.type = 'button';
                tetrahedronModeCycleBtn.className = 'card-cycle-btn';
                tetrahedronModeCycleBtn.dataset.cycleTetrahedronModeSlotId = String(slot.id);
                tetrahedronModeCycleBtn.textContent = this.isTetraBaseModeLocked(slot)
                    ? `Cycle Base (linked): ${this.getTetrahedronTriangleModeLabel(slot.params.baseTriangleMode)}`
                    : `Cycle Base: ${this.getTetrahedronTriangleModeLabel(slot.params.baseTriangleMode)}`;
                if (this.isTetraBaseModeLocked(slot)) {
                    const linkedController = this.getLinkedTetraBaseController(slot);
                    tetrahedronModeCycleBtn.title = linkedController?.primitive === 'tetrahedron'
                        ? 'Linked to attached tetrahedron (click to cycle host tetra type)'
                        : 'Linked to attached prism (click to cycle prism type)';
                }
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
            this.refreshDerivedPoints();
            this.rebuildConstructions();
            this.buildPointMarkers();
            this.updatePanelCopy();
            this.renderPointsList();
            this.renderSelectionSummary();
            this.renderActions();
            this.updateCanvasEmptyState();
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
                    const hostCenterToFace = hostFaceCenter.clone().sub(hostGroupP);
                    const hostFaceNormal = this.resolveFaceNormal(entry.faceDef, entry.slot.params)
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ);
                    const hostFaceU = this.resolveFaceUAxis(entry.faceDef, entry.slot.params)
                        .clone()
                        .applyQuaternion(hostOrientQ)
                        .applyQuaternion(hostGroupQ)
                        .normalize();

                    // Guard against inverted face metadata by forcing normals to point away from the host center.
                    if (hostCenterToFace.lengthSq() > 1e-8 && hostFaceNormal.dot(hostCenterToFace) < 0) {
                        hostFaceNormal.multiplyScalar(-1);
                        hostFaceU.multiplyScalar(-1);
                    }

                    this.applySlotTransform(def.group, slot, hostFaceCenter, hostFaceNormal, hostFaceU, {
                        hostSlot: entry.slot,
                        hostFaceDef: entry.faceDef,
                        hostGroupQuaternion: hostGroupQ.clone(),
                        hostGroupPosition: hostGroupP.clone()
                    });

                    const guestFaceDef = this.getGuestAttachFaceDef(slot, this.resolveFaceNormal(entry.faceDef, entry.slot.params), entry.faceDef.type);
                    if (guestFaceDef) {
                        this.addLinkages(
                            { slotId: entry.slotId, dims: this.resolveFaceDims(entry.faceDef, entry.slot.params) },
                            { slotId: slot.id, dims: this.getAttachmentDimsForHost(slot, guestFaceDef, entry.faceDef) }
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
        this.refreshDerivedPoints();
        this.rebuildConstructions();
        this.buildPointMarkers();
        this.updatePanelCopy();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
        if (fitCamera) {
            this.fitCameraToObject(this.compositeGroup);
        }
        this.updateCanvasEmptyState();
    }

    updateCanvasEmptyState() {
        if (!this.canvasEmptyStateEl) return;
        const shouldShow = this.compositeSlots.length === 0;
        this.canvasEmptyStateEl.classList.toggle('show', shouldShow);
    }

    getPrimitiveLocalPointMap(primitiveKey, params) {
        if (primitiveKey === 'right-triangle-prism') {
            const zFront = params.length / 2;
            const zBack = -params.length / 2;
            const [posA, posB, posC] = getTriangularPrismProfilePoints(params, zFront);
            return {
                A: posA,
                B: posB,
                C: posC,
                D: new THREE.Vector3(posA.x, posA.y, zBack),
                E: new THREE.Vector3(posB.x, posB.y, zBack),
                F: new THREE.Vector3(posC.x, posC.y, zBack)
            };
        }

        if (primitiveKey === 'tetrahedron') {
            const yBase = -params.height / 2;
            const yApex = params.height / 2;
            const [baseA, baseB, baseC] = getTetrahedronBasePoints(params, yBase);
            const apexTargetKey = params.apexPosition || 'A';
            const apexTargets = { A: baseA, B: baseB, C: baseC };
            const apexAnchor = apexTargets[apexTargetKey] || baseA;
            return {
                A: baseA,
                B: baseB,
                C: baseC,
                D: new THREE.Vector3(apexAnchor.x, yApex, apexAnchor.z)
            };
        }

        return null;
    }

    getAttachmentFaceLocalVertices(slot, faceDef) {
        if (!slot || !faceDef || !Array.isArray(faceDef.vertexIds) || faceDef.vertexIds.length !== 3) {
            return null;
        }

        const pointMap = this.getPrimitiveLocalPointMap(slot.primitive, slot.params);
        if (!pointMap) {
            return null;
        }

        const vertices = faceDef.vertexIds.map((id) => pointMap[id]?.clone()).filter(Boolean);
        return vertices.length === 3 ? vertices : null;
    }

    orientTriangleVerticesToNormal(vertices, targetNormal) {
        if (!Array.isArray(vertices) || vertices.length !== 3) {
            return null;
        }

        const oriented = vertices.map((vertex) => vertex.clone());
        const currentNormal = new THREE.Vector3()
            .crossVectors(
                oriented[1].clone().sub(oriented[0]),
                oriented[2].clone().sub(oriented[0])
            )
            .normalize();

        if (currentNormal.dot(targetNormal) < 0) {
            [oriented[1], oriented[2]] = [oriented[2], oriented[1]];
        }

        return oriented;
    }

    buildTriangleBasis(vertices) {
        const origin = vertices[0].clone();
        const u = vertices[1].clone().sub(vertices[0]).normalize();
        const n = new THREE.Vector3().crossVectors(
            vertices[1].clone().sub(vertices[0]),
            vertices[2].clone().sub(vertices[0])
        ).normalize();
        const v = new THREE.Vector3().crossVectors(n, u).normalize();
        const basis = new THREE.Matrix4().makeBasis(u, v, n);
        return { origin, basis };
    }

    getTriangleVertexSignatures(vertices) {
        return vertices.map((vertex, index) => {
            const distances = vertices
                .filter((_, otherIndex) => otherIndex !== index)
                .map((other) => vertex.distanceToSquared(other))
                .sort((left, right) => left - right);
            return distances;
        });
    }

    chooseBestTriangleVertexAlignment(guestVertices, hostVertices, targetNormal) {
        const guestSignatures = this.getTriangleVertexSignatures(guestVertices);
        const permutations = [
            [0, 1, 2],
            [0, 2, 1],
            [1, 2, 0],
            [1, 0, 2],
            [2, 0, 1],
            [2, 1, 0]
        ];

        let bestVertices = hostVertices;
        let bestScore = Number.POSITIVE_INFINITY;

        permutations.forEach((permutation) => {
            const candidate = permutation.map((index) => hostVertices[index]);
            const candidateNormal = new THREE.Vector3()
                .crossVectors(
                    candidate[1].clone().sub(candidate[0]),
                    candidate[2].clone().sub(candidate[0])
                )
                .normalize();
            if (candidateNormal.dot(targetNormal) <= 1e-8) {
                return;
            }

            const candidateSignatures = this.getTriangleVertexSignatures(candidate);
            const score = guestSignatures.reduce((total, signature, index) => {
                return total
                    + Math.abs(signature[0] - candidateSignatures[index][0])
                    + Math.abs(signature[1] - candidateSignatures[index][1]);
            }, 0);

            if (score < bestScore) {
                bestScore = score;
                bestVertices = candidate;
            }
        });

        return bestVertices.map((vertex) => vertex.clone());
    }

    computeBestTriangleAttachmentTransform(guestVertices, hostVertices, targetGuestNormal, isCandidateValid = null) {
        const guestOrders = [
            [0, 1, 2],
            [0, 2, 1]
        ];
        const hostOrders = [
            [0, 1, 2], [0, 2, 1],
            [1, 0, 2], [1, 2, 0],
            [2, 0, 1], [2, 1, 0]
        ];

        let best = null;

        guestOrders.forEach((guestOrder) => {
            const src = guestOrder.map((index) => guestVertices[index].clone());
            const srcBasis = this.buildTriangleBasis(src);
            const srcNormal = new THREE.Vector3()
                .crossVectors(
                    src[1].clone().sub(src[0]),
                    src[2].clone().sub(src[0])
                )
                .normalize();

            hostOrders.forEach((hostOrder) => {
                const dst = hostOrder.map((index) => hostVertices[index].clone());
                const dstBasis = this.buildTriangleBasis(dst);

                const transformMatrix = dstBasis.basis.clone().multiply(srcBasis.basis.clone().invert());
                const quaternion = new THREE.Quaternion().setFromRotationMatrix(transformMatrix);
                const position = dstBasis.origin.clone().sub(srcBasis.origin.clone().applyQuaternion(quaternion));

                const rotatedNormal = srcNormal.clone().applyQuaternion(quaternion).normalize();
                if (rotatedNormal.dot(targetGuestNormal) < 0.999) {
                    return;
                }

                if (typeof isCandidateValid === 'function' && !isCandidateValid(quaternion, position, dst)) {
                    return;
                }

                const error = src.reduce((total, vertex, index) => {
                    const transformed = vertex.clone().applyQuaternion(quaternion).add(position);
                    return total + transformed.distanceToSquared(dst[index]);
                }, 0);

                if (!best || error < best.error) {
                    best = { quaternion, position, error };
                }
            });
        });

        return best;
    }

    tryApplyExactTriangleFaceTransform(slotGroup, slot, hostFaceNormal, options = {}) {
        const { hostSlot, hostFaceDef, hostGroupQuaternion, hostGroupPosition } = options;
        if (!hostSlot || !hostFaceDef || hostFaceDef.type !== 'triangle') {
            return false;
        }

        const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal, hostFaceDef.type);
        if (!guestFaceDef || guestFaceDef.type !== 'triangle') {
            return false;
        }

        const guestLocalVertices = this.getAttachmentFaceLocalVertices(slot, guestFaceDef);
        const hostLocalVertices = this.getAttachmentFaceLocalVertices(hostSlot, hostFaceDef);
        if (!guestLocalVertices || !hostLocalVertices) {
            return false;
        }

        const guestOrientQ = this.getOrientationQuaternion(slot.primitive, slot.orientation);
        const hostOrientQ = this.getOrientationQuaternion(hostSlot.primitive, hostSlot.orientation);
        const targetGuestNormal = hostFaceNormal.clone().negate().normalize();

        const guestOrientedVertices = guestLocalVertices.map((vertex) => vertex.applyQuaternion(guestOrientQ));
        const hostWorldVertices = hostLocalVertices.map((vertex) => vertex.applyQuaternion(hostOrientQ).applyQuaternion(hostGroupQuaternion).add(hostGroupPosition));

        const pointMap = this.getPrimitiveLocalPointMap(slot.primitive, slot.params);
        const apexLocal = slot.primitive === 'tetrahedron' ? pointMap?.D?.clone().applyQuaternion(guestOrientQ) : null;

        const bestTransform = this.computeBestTriangleAttachmentTransform(
            guestOrientedVertices,
            hostWorldVertices,
            targetGuestNormal,
            (quaternion, position, dstVertices) => {
                if (!apexLocal) {
                    return true;
                }

                const transformedApex = apexLocal.clone().applyQuaternion(quaternion).add(position);
                const facePoint = dstVertices[0];
                const signedDistance = transformedApex.clone().sub(facePoint).dot(hostFaceNormal.clone().normalize());
                return signedDistance > 1e-5;
            }
        );
        if (!bestTransform) {
            return false;
        }

        slotGroup.quaternion.copy(bestTransform.quaternion);
        slotGroup.position.copy(bestTransform.position);
        return true;
    }

    applySlotTransform(slotGroup, slot, hostFaceCenter, hostFaceNormal, hostFaceUWorld = null, options = {}) {
        if (this.tryApplyExactTriangleFaceTransform(slotGroup, slot, hostFaceNormal, options)) {
            return;
        }

        const { hostFaceDef: _hostFaceDef } = options;
        const guestFaceDef = this.getGuestAttachFaceDef(slot, hostFaceNormal, _hostFaceDef?.type ?? null);
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

        const quarterTurns = this.getAttachmentQuarterTurns(slot, guestFaceDef, _hostFaceDef);
        if (quarterTurns !== 0) {
            const extraTwist = new THREE.Quaternion().setFromAxisAngle(targetGuestNormal, quarterTurns * (Math.PI / 2));
            Q.premultiply(extraTwist);
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

            const preferredLabel = this.baseLabelOverrides.get(point.id) || point.label;
            const label = this.getUniqueDisplayLabel(preferredLabel, usedLabels);
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
            sphere: [],
            hemisphere: [],
            cylinder: [],
            cone: [],
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

    hasExplicitSceneSegmentBetween(pointIds) {
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
            return !!definition
                && definition.kind === 'segment'
                && definition.hidden !== true
                && Array.isArray(definition.pointIds)
                && definition.pointIds.length === 2
                && matchesPair(definition.pointIds);
        });
    }

    canAttachLabelToPointPair(pointIds) {
        // Labels can attach to any valid edge connection, including raw primitive edges.
        return this.hasPrimitiveEdgeBetween(pointIds)
            || this.hasSceneSegmentBetween(pointIds)
            || this.hasSharedDerivedEdgeParent(pointIds);
    }

    canAttachMidpointToPointPair(pointIds) {
        return this.hasSegmentLikeConnection(pointIds);
    }

    hasSegmentLikeConnection(pointIds) {
        return this.hasPrimitiveEdgeBetween(pointIds)
            || this.hasSceneSegmentBetween(pointIds)
            || this.hasSharedDerivedEdgeParent(pointIds);
    }

    hasSharedDerivedEdgeParent(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 2 || pointIds[0] === pointIds[1]) {
            return false;
        }

        const firstPair = this.getDerivedEdgeBasePairForPointId(pointIds[0]);
        const secondPair = this.getDerivedEdgeBasePairForPointId(pointIds[1]);
        if (!firstPair || !secondPair) {
            return false;
        }

        return firstPair[0] === secondPair[0] && firstPair[1] === secondPair[1];
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

    findAngleObject(pointIds) {
        if (!Array.isArray(pointIds) || pointIds.length !== 3) {
            return null;
        }

        const vertexId = pointIds[1];
        const endpointPair = this.normalizePointPairIds([pointIds[0], pointIds[2]]);
        if (!vertexId || !endpointPair) {
            return null;
        }

        return this.sceneObjects.find((entry) => {
            const definition = entry.definition;
            if (!definition || definition.kind !== 'angle' || !Array.isArray(definition.pointIds) || definition.pointIds.length !== 3) {
                return false;
            }

            if (definition.pointIds[1] !== vertexId) {
                return false;
            }

            const candidatePair = this.normalizePointPairIds([definition.pointIds[0], definition.pointIds[2]]);
            return !!candidatePair && candidatePair[0] === endpointPair[0] && candidatePair[1] === endpointPair[1];
        }) || null;
    }

    findAngleLabelObjectsForPolygon(pointIds) {
        return this.sceneObjects.filter((entry) => {
            const def = entry.definition;
            if (!def || def.kind !== 'angle' || !Array.isArray(def.pointIds) || def.pointIds.length !== 3) return false;
            return def.pointIds.every((id) => pointIds.includes(id));
        });
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

    getRatioSignatureForPointId(pointId) {
        const point = this.getPointById(pointId);
        if (!point || !point.isDerived) {
            return null;
        }

        if (typeof point.signature === 'string' && point.signature.startsWith('ratio|')) {
            return point.signature;
        }

        if (typeof point.id === 'string' && point.id.startsWith('derived-ratio|')) {
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

    findRatioPointObjectByPointId(pointId) {
        const signature = this.getRatioSignatureForPointId(pointId);
        if (!signature) {
            return null;
        }

        return this.sceneObjects.find((entry) => entry.definition?.kind === 'ratio-point' && entry.definition.signature === signature) || null;
    }

    findDerivedEdgePointObjectByPointId(pointId) {
        return this.findMidpointObjectByPointId(pointId) || this.findRatioPointObjectByPointId(pointId);
    }

    hasRatioPointForSignature(signature) {
        if (!signature) {
            return false;
        }

        return this.sceneObjects.some((entry) => entry.definition?.kind === 'ratio-point' && entry.definition.signature === signature);
    }

    removeEdgeLabelsForPointPair(pointIds, options = {}) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return;
        }

        const onlyIfDisconnected = options.onlyIfDisconnected === true;
        if (onlyIfDisconnected && this.canAttachLabelToPointPair(normalized)) {
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
        if (onlyIfDisconnected && this.canAttachMidpointToPointPair(normalized)) {
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

    removeRatioPointsForPair(pointIds, options = {}) {
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) {
            return;
        }

        const onlyIfDisconnected = options.onlyIfDisconnected === true;
        if (onlyIfDisconnected && this.canAttachMidpointToPointPair(normalized)) {
            return;
        }

        const survivors = [];
        this.sceneObjects.forEach((entry) => {
            if (entry.definition?.kind !== 'ratio-point') {
                survivors.push(entry);
                return;
            }

            const pair = this.normalizePointPairIds(entry.definition.pointIds || []);
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

    getPolygonEdgePairsInOrder(pointIds) {
        if (!Array.isArray(pointIds)) {
            return [];
        }

        if (pointIds.length === 3) {
            return this.getTriangleEdgePairs(pointIds);
        }

        if (pointIds.length === 4) {
            return this.getPlaneEdgePairs(pointIds);
        }

        return [];
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
            this.removeEdgeLabelsForPointPair(pair, { onlyIfDisconnected: true });
            this.removeMidpointPointsForPair(pair, { onlyIfDisconnected: true });
            this.removeRatioPointsForPair(pair, { onlyIfDisconnected: true });
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

    makeRatioPointSignature(pointIds, leftValue, rightValue) {
        const normalizedPair = this.normalizePointPairIds(pointIds);
        const ratio = this.reduceRatio(leftValue, rightValue);
        if (!normalizedPair || !ratio) {
            return null;
        }

        const usesNormalizedOrder = normalizedPair[0] === pointIds[0] && normalizedPair[1] === pointIds[1];
        const left = usesNormalizedOrder ? ratio.left : ratio.right;
        const right = usesNormalizedOrder ? ratio.right : ratio.left;
        return `ratio|${normalizedPair[0]}|${normalizedPair[1]}|${left}|${right}`;
    }

    getDerivedEdgeBasePairFromSignature(signature) {
        if (typeof signature !== 'string') {
            return null;
        }

        const midpointMatch = /^midpoint\|([^|]+)\|([^|]+)$/.exec(signature);
        if (midpointMatch) {
            return this.normalizePointPairIds([midpointMatch[1], midpointMatch[2]]);
        }

        const ratioMatch = /^ratio\|([^|]+)\|([^|]+)\|([1-9]\d*)\|([1-9]\d*)$/.exec(signature);
        if (ratioMatch) {
            return this.normalizePointPairIds([ratioMatch[1], ratioMatch[2]]);
        }

        return null;
    }

    getDerivedEdgeBasePairForPointId(pointId) {
        const midpointSignature = this.getMidpointSignatureForPointId(pointId);
        if (midpointSignature) {
            return this.getDerivedEdgeBasePairFromSignature(midpointSignature);
        }

        const ratioSignature = this.getRatioSignatureForPointId(pointId);
        if (ratioSignature) {
            return this.getDerivedEdgeBasePairFromSignature(ratioSignature);
        }

        return null;
    }

    async changeSelectedPointLabel() {
        const pointId = this.selectedPoints[0];
        const point = this.getPointById(pointId);
        if (!point) {
            return;
        }

        const nextLabelRaw = await this.showPromptModal(`Change label for point ${point.label}`, point.label);
        if (nextLabelRaw == null) {
            return;
        }

        const nextLabel = this.normalizePointLabelInput(nextLabelRaw);
        if (!nextLabel) {
            await this.showAlertModal('Use label format A or A1 (single letter, optional single digit).');
            return;
        }

        if (!this.isPointLabelAvailable(nextLabel, point.id)) {
            await this.showAlertModal(`Label ${nextLabel} is already in use.`);
            return;
        }

        if (point.isDerived) {
            point.label = nextLabel;
            if (point.signature) {
                this.derivedLabelOverrides.set(point.signature, nextLabel);
            }
        } else {
            point.label = nextLabel;
            if (Array.isArray(point.sourceIds) && point.sourceIds.length > 0) {
                point.sourceIds.forEach((sourceId) => {
                    this.baseLabelOverrides.set(sourceId, nextLabel);
                });
            }
            this.refreshDerivedPoints();
        }

        this.buildPointMarkers();
        this.renderObjectsList();
        this.clearSelection();
    }

    updatePanelCopy() {
        if (!this.primitiveChip || !this.orientationChip) {
            return;
        }

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

            const centersMode = this.getCuboidFaceCentersMode(params);
            if (centersMode === 'front-back' || centersMode === 'all') {
                points.push(
                    { id: 'I', label: 'I', description: 'front centre', position: new THREE.Vector3(0, 0, depth / 2) },
                    { id: 'J', label: 'J', description: 'back centre', position: new THREE.Vector3(0, 0, -depth / 2) }
                );
            }
            if (centersMode === 'left-right' || centersMode === 'all') {
                points.push(
                    { id: 'K', label: 'K', description: 'left centre', position: new THREE.Vector3(-width / 2, 0, 0) },
                    { id: 'L', label: 'L', description: 'right centre', position: new THREE.Vector3(width / 2, 0, 0) }
                );
            }
            if (centersMode === 'top-bottom' || centersMode === 'all') {
                points.push(
                    { id: 'M', label: 'M', description: 'top centre', position: new THREE.Vector3(0, height / 2, 0) },
                    { id: 'N', label: 'N', description: 'bottom centre', position: new THREE.Vector3(0, -height / 2, 0) }
                );
            }

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
            opacity: this.ghostFaces ? 0.14 : 0.46,
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
            new THREE.LineBasicMaterial({ color: this.getEdgeColor(), transparent: false, opacity: 1 })
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
                    true,
                    this.getEdgeColor()
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
                    new THREE.LineBasicMaterial({ color: this.getEdgeColor(), transparent: false, opacity: 1 })
                );
                circle.renderOrder = 7;
                group.add(circle);
            });
        }

        return { group, mesh, points, boundsRadius };
    }

    updatePrimitiveMaterial() {
        this.primitiveMeshes.forEach((mesh) => {
            mesh.material.opacity = this.ghostFaces ? 0.14 : 0.46;
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

        const markerRadius = this.displaySizeMode === 'small' ? 0.075 : 0.1;
        const markerGeometry = new THREE.SphereGeometry(markerRadius, 18, 18);
        this.getAllPoints().forEach((point) => {
            const markerColor = point.isDerived ? 0x2e7d32 : this.getEdgeColor();
            const marker = new THREE.Mesh(markerGeometry, new THREE.MeshBasicMaterial({ color: markerColor }));
            marker.position.copy(point.position);
            marker.visible = this.pointMarkersVisible;
            this.primitiveGroup.add(marker);
            this.pointMarkers.set(point.id, marker);

            const labelBackground = point.isDerived ? '#d9f9d6' : '#ffd84d';
            const sprite = this.createTextSprite(point.label, {
                fontSize: 52,
                textColor: this.getLabelTextColor(),
                background: labelBackground,
                borderColor: '#000000'
            });
            sprite.position.copy(point.position.clone().add(new THREE.Vector3(0.18, 0.22, 0.18)));
            this.primitiveGroup.add(sprite);
            this.pointSprites.push(sprite);
        });
        this.syncAllLabelVisibility();
    }

    renderPointsList() {
        this.pointsListEl.innerHTML = '';
        const allPoints = this.getAllPoints();
        allPoints.forEach((point) => {
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
        if (allPoints.length > 0 && this.selectedPoints.length === 0 && !this.pointsHintDismissed) {
            const hint = document.createElement('p');
            hint.className = 'section-note points-select-hint';
            hint.textContent = 'Select 1–4 labelled points in order to reveal available actions.';
            this.pointsListEl.appendChild(hint);
        }
        this.updatePrimarySectionCounts();
        this.updatePointMarkerStyles();
    }

    updatePrimarySectionCounts() {
        const updateSectionTitleCount = (titleEl, baseTitle, count) => {
            if (!titleEl) return;
            titleEl.textContent = baseTitle;
            if (count > 0) {
                const countEl = document.createElement('span');
                countEl.className = 'section-object-count';
                countEl.textContent = `${count}`;
                titleEl.appendChild(countEl);
            }
        };

        updateSectionTitleCount(this.primitiveSectionTitle, this.primitiveSectionBaseTitle, this.compositeSlots.length);
        updateSectionTitleCount(this.pointsSectionTitle, this.pointsSectionBaseTitle, this.getAllPoints().length);
    }

    updatePointMarkerStyles() {
        this.getAllPoints().forEach((point) => {
            const marker = this.pointMarkers.get(point.id);
            if (!marker) return;

            const isSelected = this.selectedPoints.includes(point.id);
            const baseColor = point.isDerived ? 0x2e7d32 : this.getEdgeColor();
            marker.material.color.set(isSelected ? 0x4a90e2 : baseColor);
            marker.scale.setScalar(isSelected ? 1.45 : 1);
        });
    }

    togglePointSelection(pointId) {
        this.pointsHintDismissed = true;
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
            const derivedPointObject = this.findDerivedEdgePointObjectByPointId(this.selectedPoints[0]);
            if (derivedPointObject) {
                const label = derivedPointObject.definition?.kind === 'ratio-point' ? 'Delete Ratio Point' : 'Delete Midpoint';
                baseActions.push({ key: 'delete-derived-point', label });
            }
        }

        if (this.selectedPoints.length === 2) {
            const hasExistingSegment = this.hasExplicitSceneSegmentBetween(this.selectedPoints) || this.hasPrimitiveEdgeBetween(this.selectedPoints);
            const filteredActions = baseActions.filter((action) => action.key !== 'segment' || !hasExistingSegment);

            if (this.canAttachLabelToPointPair(this.selectedPoints)) {
                const hasExistingLabel = !!this.findEdgeLabelObject(this.selectedPoints);
                filteredActions.push({ key: 'edge-label', label: hasExistingLabel ? 'Change Label' : 'Add Label' });
            }

            if (this.canAttachMidpointToPointPair(this.selectedPoints) && !this.hasMidpointForPair(this.selectedPoints)) {
                filteredActions.push({ key: 'add-midpoint', label: 'Add Midpoint' });
            }

            if (this.canAttachMidpointToPointPair(this.selectedPoints)) {
                filteredActions.push({ key: 'add-ratio-point', label: 'Add Ratio Point' });
            }

            return filteredActions;
        }

        if (this.selectedPoints.length === 3) {
            return baseActions
                .filter((action) => {
                    if (action.key !== 'angle') {
                        return true;
                    }
                    return this.canAttachAngleFromOrderedPoints(this.selectedPoints);
                })
                .map((action) => {
                    if (action.key !== 'angle') {
                        return action;
                    }
                    const hasExistingAngle = !!this.findAngleObject(this.selectedPoints);
                    return {
                        ...action,
                        label: hasExistingAngle ? 'Change Angle' : 'Add Angle'
                    };
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
        const edgeDivisionDefinitions = this.sceneObjects
            .map((entry) => entry.definition)
            .filter((definition) => (definition?.kind === 'midpoint-point' || definition?.kind === 'ratio-point') && Array.isArray(definition.pointIds) && definition.pointIds.length === 2);
        const seenEdgeDivisionSignatures = new Set();

        edgeDivisionDefinitions.forEach((definition) => {
            if (!this.canAttachMidpointToPointPair(definition.pointIds)) {
                return;
            }

            const start = priorPointMap.get(definition.pointIds[0]);
            const end = priorPointMap.get(definition.pointIds[1]);
            if (!start || !end) {
                return;
            }

            let signature = null;
            let position = null;
            let description = 'derived point';

            if (definition.kind === 'midpoint-point') {
                signature = definition.signature || this.makeMidpointSignature(definition.pointIds);
                position = start.clone().lerp(end, 0.5);
                description = 'derived midpoint';
            }

            if (definition.kind === 'ratio-point') {
                const ratio = this.reduceRatio(definition.ratioA, definition.ratioB);
                if (!ratio || ratio.left === ratio.right) {
                    return;
                }

                signature = definition.signature || this.makeRatioPointSignature(definition.pointIds, ratio.left, ratio.right);
                const distanceRatio = ratio.left / (ratio.left + ratio.right);
                position = start.clone().lerp(end, distanceRatio);
                description = `derived point in ratio ${ratio.left}:${ratio.right}`;
            }

            if (!signature || !position || seenEdgeDivisionSignatures.has(signature) || existingPointNear(position)) {
                return;
            }

            seenEdgeDivisionSignatures.add(signature);
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
            basePoints.set(id, position.clone());
            derived.push({
                id,
                label,
                signature,
                description,
                position,
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

        if (!this.isRestoringSharedState) {
            Array.from(this.derivedLabelOverrides.keys()).forEach((signature) => {
                if (!activeDerivedSignatures.has(signature)) {
                    this.derivedLabelOverrides.delete(signature);
                }
            });
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

    async runAction(actionKey) {
        const vectors = this.getSelectedVectors();
        if (vectors.some((value) => !value)) {
            return;
        }
        
        const selectedLabels = this.selectedPoints.map((pointId) => this.getPointById(pointId)?.label || pointId);

        if (actionKey === 'change-point-label') {
            await this.changeSelectedPointLabel();
            this.closePanelOnMobile();
            return;
        }

        if (actionKey === 'delete-derived-point') {
            if (this.selectedPoints.length !== 1) {
                return;
            }

            const derivedPointObject = this.findDerivedEdgePointObjectByPointId(this.selectedPoints[0]);
            if (!derivedPointObject) {
                return;
            }

            this.deleteObject(derivedPointObject.id);
            this.clearSelection();
            this.closePanelOnMobile();
            return;
        }

        if (actionKey === 'segment') {
            const ids = [...this.selectedPoints];
            if (this.hasExplicitSceneSegmentBetween(ids)) {
                this.clearSelection();
                this.closePanelOnMobile();
                return;
            }

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
            const nextText = await this.showPromptModal(promptText, currentText, { quickSymbols: true });
            if (nextText == null || !nextText.trim()) {
                return;
            }

            const normalizedText = nextText.trim();
            const midpoint = vectors[0].clone().lerp(vectors[1], 0.5).add(new THREE.Vector3(0.2, 0.2, 0.2));
            const sprite = this.createTextSprite(normalizedText, {
                fontSize: 46,
                textColor: this.getLabelTextColor(),
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
                this.focusObjectSectionForType(existingLabel.type, existingLabel.definition);
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
            if (!ids || !this.canAttachMidpointToPointPair(ids) || this.hasMidpointForPair(ids)) {
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

        if (actionKey === 'add-ratio-point') {
            const orderedIds = [...this.selectedPoints];
            if (orderedIds.length !== 2 || !this.canAttachMidpointToPointPair(orderedIds)) {
                return;
            }

            const ratio = await this.showRatioPromptModal(`Split ${selectedLabels[0]} : ${selectedLabels[1]} in ratio`, 1, 2);
            if (!ratio) {
                return;
            }

            if (ratio.left === ratio.right) {
                await this.showAlertModal('Use Add Midpoint for a 1:1 split.');
                return;
            }

            const signature = this.makeRatioPointSignature(orderedIds, ratio.left, ratio.right);
            if (!signature || this.hasRatioPointForSignature(signature)) {
                await this.showAlertModal('That ratio point already exists on this edge.');
                return;
            }

            this.addSceneObject({
                type: 'label',
                name: `Ratio Point ${this.formatPointSequence(orderedIds)}`,
                subtitle: `Ratio ${ratio.left}:${ratio.right}`,
                object3D: new THREE.Group(),
                definition: {
                    kind: 'ratio-point',
                    pointIds: orderedIds,
                    ratioA: ratio.left,
                    ratioB: ratio.right,
                    signature
                }
            });
            const derivedId = `derived-${signature}`;
            this.addSceneObject({
                type: 'segment',
                name: 'Ghost sub-segment A',
                subtitle: 'Ratio-point sub-segment',
                object3D: new THREE.Group(),
                definition: { kind: 'segment', pointIds: [orderedIds[0], derivedId], hidden: true }
            });
            this.addSceneObject({
                type: 'segment',
                name: 'Ghost sub-segment B',
                subtitle: 'Ratio-point sub-segment',
                object3D: new THREE.Group(),
                definition: { kind: 'segment', pointIds: [derivedId, orderedIds[1]], hidden: true }
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

            const existingAngle = this.findAngleObject(ids);
            const defaultAngleLabel = existingAngle?.definition?.text || this.formatPointSequence(ids);
            let angleLabelInput = defaultAngleLabel;
            const promptText = existingAngle
                ? `Change label for angle at ${selectedLabels[1]}`
                : `Label for angle at ${selectedLabels[1]}`;
            while (true) {
                const nextLabel = await this.showPromptModal(promptText, angleLabelInput, { quickSymbols: true });
                if (nextLabel == null) {
                    return;
                }

                angleLabelInput = nextLabel.trim();
                if (angleLabelInput) {
                    break;
                }

                await this.showAlertModal('Angle label cannot be blank.');
            }

            const color = existingAngle?.definition?.color || this.nextConstructionColor();
            const angleGroup = this.createAngleMarker(vectors[0], vectors[1], vectors[2], angleLabelInput, color);

            if (existingAngle) {
                existingAngle.name = `Angle ${angleLabelInput}`;
                existingAngle.subtitle = `Angle at ${selectedLabels[1]}`;
                existingAngle.definition = {
                    kind: 'angle',
                    pointIds: ids,
                    text: angleLabelInput,
                    color
                };
                this.scene.remove(existingAngle.object3D);
                this.disposeObject3D(existingAngle.object3D);
                angleGroup.userData.sceneObjectId = existingAngle.id;
                existingAngle.object3D = angleGroup;
                existingAngle.object3D.visible = existingAngle.visible;
                this.scene.add(existingAngle.object3D);
                this.focusObjectSectionForType(existingAngle.type, existingAngle.definition);
                this.renderObjectsList();
                this.renderSelectionSummary();
                this.renderActions();
            } else {
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
                name: `Quadrilateral ${this.formatPointSequence(orderedIds)}`,
                subtitle: 'Four-point coplanar patch',
                object3D: plane,
                definition: {
                    kind: 'plane',
                    pointIds: orderedIds,
                    color,
                    opacity: 0.2
                }
            });
            this.ensureHiddenSupportSegmentsForPairs(this.getPlaneEdgePairs(orderedIds), 'Quadrilateral support edge', planeObject?.id ?? null);
        }

        this.clearSelection();
        this.closePanelOnMobile();
    }

    createSegment(start, end, color) {
        const segment = this.createThickPolyline([start, end], color, 5);
        segment.renderOrder = 21;
        return segment;
    }

    createThickPolyline(points, color, width = 5) {
        const geometry = new LineGeometry();
        geometry.setPositions(points.flatMap((point) => [point.x, point.y, point.z]));

        const material = new LineMaterial({
            color,
            linewidth: width,
            worldUnits: false,
            transparent: false,
            depthTest: false,
            depthWrite: false
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
        const radius = Math.min(a.distanceTo(vertex), c.distanceTo(vertex)) * 0.22;
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

        const labelRadius = radius * 0.6;
        const labelPoint = vertex.clone()
            .add(dir1.clone().multiplyScalar(Math.cos(rawAngle / 2) * labelRadius))
            .add(tangent.clone().multiplyScalar(Math.sin(rawAngle / 2) * labelRadius));
        const label = this.createTextSprite(`${angleText}`, {
            fontSize: 42,
            textColor: this.getLabelTextColor(),
            background: '#9de7ff',
            borderColor: '#000000'
        });
        label.position.copy(labelPoint);

        const group = new THREE.Group();
        group.add(line, label);
        return group;
    }

    createTextSprite(text, options = {}) {
        const baseFontSize = options.fontSize || 48;
        const displayScale = this.displaySizeMode === 'small' ? 0.78 : 1.14;
        const fontSize = Math.max(12, Math.round(baseFontSize * displayScale));
        const badgeVisible = options.forceBadge === true || this.labelMode === 'badge';
        const paddingX = badgeVisible
            ? Math.max(8, Math.round(14 * displayScale))
            : Math.max(2, Math.round(4 * displayScale));
        const paddingY = badgeVisible
            ? Math.max(5, Math.round(9 * displayScale))
            : Math.max(2, Math.round(4 * displayScale));
        const dpr = Math.min(window.devicePixelRatio || 1, 2);
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        context.font = `700 ${fontSize}px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`;
        const metrics = context.measureText(text);
        const logicalWidth = Math.ceil(metrics.width + paddingX * 2);
        const logicalHeight = Math.ceil(fontSize + paddingY * 2);
        canvas.width = Math.ceil(logicalWidth * dpr);
        canvas.height = Math.ceil(logicalHeight * dpr);

        const drawContext = canvas.getContext('2d');
        drawContext.setTransform(dpr, 0, 0, dpr, 0, 0);
        drawContext.clearRect(0, 0, logicalWidth, logicalHeight);
        drawContext.font = `700 ${fontSize}px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`;
        drawContext.textBaseline = 'middle';
        drawContext.textAlign = 'center';
        drawContext.lineJoin = 'round';
        drawContext.lineCap = 'round';

        const background = options.background || '#ffd84d';
        const borderColor = options.borderColor || '#000000';
        const borderWidth = badgeVisible
            ? Math.max(2, Math.round((options.borderWidth || 4) * displayScale))
            : 0;
        const radius = badgeVisible
            ? Math.max(8, Math.round((options.cornerRadius || 16) * displayScale))
            : 0;

        if (badgeVisible) {
            drawContext.fillStyle = background;
            this.drawRoundedRect(drawContext, borderWidth / 2, borderWidth / 2, logicalWidth - borderWidth, logicalHeight - borderWidth, radius);
            drawContext.fill();

            drawContext.lineWidth = borderWidth;
            drawContext.strokeStyle = borderColor;
            drawContext.stroke();
        }

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

    setObjectSectionCollapsedState(sectionKey, collapsed) {
        const section = this.objectSections[sectionKey];
        if (!section) {
            return;
        }

        this.objectGroupCollapsed[sectionKey] = collapsed;
        section.content.classList.toggle('collapsed', collapsed);
        section.header.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        section.arrow.textContent = collapsed ? '\u25B6\uFE0E' : '\u25BC\uFE0E';
    }

    focusObjectSectionForType(type, definition = null) {
        const sectionByType = {
            triangle: 'triangles',
            segment: 'segments',
            angle: 'angles',
            plane: 'planes',
            label: 'labels'
        };

        const sectionKey = sectionByType[type];
        if (!sectionKey) {
            return;
        }

        if (definition?.hidden || definition?.kind === 'midpoint-point' || definition?.kind === 'ratio-point') {
            return;
        }

        Object.keys(this.objectSections).forEach((key) => {
            this.setObjectSectionCollapsedState(key, key !== sectionKey);
        });
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
        this.focusObjectSectionForType(type, definition);
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
        this.syncEdgeLabelVisibility();
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
                textColor: this.getLabelTextColor(),
                background: '#b9f18a',
                borderColor: '#000000'
            });
            const midpoint = vectors[0].clone().lerp(vectors[1], 0.5).add(new THREE.Vector3(0.2, 0.2, 0.2));
            sprite.position.copy(midpoint);
            return sprite;
        }

        if (definition.kind === 'midpoint-point' || definition.kind === 'ratio-point') {
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
                if (item.definition?.kind === 'midpoint-point' || item.definition?.kind === 'ratio-point') return false;
                return true;
            });
            sec.list.innerHTML = '';
            items.forEach((item) => sec.list.appendChild(this.renderObjectItem(item)));
            if (sec.title) {
                sec.title.textContent = sec.baseTitle;
                if (items.length > 0) {
                    const countEl = document.createElement('span');
                    countEl.className = 'section-object-count';
                    countEl.textContent = `${items.length}`;
                    sec.title.appendChild(countEl);
                }
            }
        }
    }



    escapeHtmlAttribute(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    renderObjectItem(item) {
        const row = document.createElement('div');
        row.className = item.visible ? 'object-item' : 'object-item disabled';
        const itemColor = item.definition?.color != null
            ? `#${item.definition.color.toString(16).padStart(6, '0')}`
            : null;
        const visibilityColor = item.type === 'label' ? '#b9f18a' : itemColor;
        const displayName = this.getSceneObjectDisplayName(item);
        const isAngle = item.type === 'angle' && item.definition?.kind === 'angle' && Array.isArray(item.definition?.pointIds) && item.definition.pointIds.length === 3;
        const displayNameHtml = isAngle ? this.getAngleDisplayHtml(item) : displayName;
        const showSubtitle = item.type === 'label';
        const subtitleText = typeof item.subtitle === 'string' ? item.subtitle : '';
        const isEditableLabel = item.type === 'label'
            && (item.definition?.kind === 'edge-label' || item.definition?.kind === 'length-label' || item.definition?.kind === 'point-label');
        const currentLabelText = isEditableLabel ? String(item.definition?.text ?? '').trim() : '';
        const labelButtonTooltip = currentLabelText ? `Change label: ${currentLabelText}` : 'Change label text';
        const subtitleHtml = isAngle
            ? `<button type="button" class="object-label-edit" data-edit-angle-object-id="${item.id}" title="Change angle label">Change Angle</button>`
            : (isEditableLabel
                ? `<button type="button" class="object-label-edit" data-edit-label-object-id="${item.id}" aria-label="${this.escapeHtmlAttribute(labelButtonTooltip)}" title="${this.escapeHtmlAttribute(labelButtonTooltip)}">Change Label</button>`
                : (showSubtitle && subtitleText ? `<span>${subtitleText}</span>` : ''));
        const isInspectablePolygon = (item.type === 'triangle' && item.definition?.pointIds?.length === 3)
            || (item.type === 'plane' && item.definition?.pointIds?.length === 4);
        const inspectObjectName = item.type === 'plane' ? 'quadrilateral' : 'triangle';
        const extractButtonHtml = isInspectablePolygon
            ? `<button type="button" class="object-extract" data-extract-object-id="${item.id}" aria-label="Inspect ${inspectObjectName} in a 2D view" title="Open this ${inspectObjectName} in a 2D teaching view showing its true shape, edge labels, and angle markers"${item.visible ? '' : ' disabled'}>Inspect</button>`
            : '';
        if (item.type === 'label') {
            row.style.borderLeftColor = '#b9f18a';
        }
        if (itemColor) {
            row.style.borderLeftColor = itemColor;
        }
        row.innerHTML = `
            <div class="object-name">
                <strong>${displayNameHtml}</strong>
                ${subtitleHtml}
                ${extractButtonHtml}
            </div>
            <div class="object-controls">
                <button
                    type="button"
                    class="object-visibility-btn"
                    data-toggle-object-id="${item.id}"
                    aria-label="${item.visible ? 'Hide' : 'Show'} object"
                    title="Click to ${item.visible ? 'hide' : 'show'} object"
                    style="background-color: ${item.visible && visibilityColor ? visibilityColor : 'transparent'};"
                ></button>
                <button type="button" class="object-delete" data-delete-object-id="${item.id}" aria-label="Delete object" title="Delete object">X</button>
            </div>
        `;
        return row;
    }

    async editLabelFromObjectCard(objectId) {
        const item = this.sceneObjects.find((entry) => entry.id === objectId && entry.type === 'label');
        if (!item || !item.definition) {
            return;
        }

        const def = item.definition;
        const currentText = def.text || '';
        let promptText = 'Change label text';

        if ((def.kind === 'edge-label' || def.kind === 'length-label') && Array.isArray(def.pointIds) && def.pointIds.length === 2) {
            promptText = `Change label for ${this.formatPointSequence(def.pointIds)}`;
        }
        if (def.kind === 'point-label') {
            const point = this.getPointById(def.pointId);
            promptText = `Change label for point ${point?.label || def.pointId || ''}`.trim();
        }

        const nextText = await this.showPromptModal(promptText, currentText, { quickSymbols: true });
        if (nextText == null || !nextText.trim()) {
            return;
        }

        const normalizedText = nextText.trim();
        item.definition = {
            ...def,
            text: normalizedText
        };
        item.subtitle = normalizedText;

        const rebuiltObject = this.createObjectFromDefinition(item.definition);
        if (!rebuiltObject) {
            return;
        }

        if (item.object3D) {
            this.scene.remove(item.object3D);
            this.disposeObject3D(item.object3D);
        }
        rebuiltObject.userData.sceneObjectId = item.id;
        rebuiltObject.visible = item.visible;
        item.object3D = rebuiltObject;
        this.scene.add(rebuiltObject);

        this.focusObjectSectionForType(item.type, item.definition);
        this.renderObjectsList();
        this.syncEdgeLabelVisibility();
    }

    async editAngleFromObjectCard(objectId) {
        const item = this.sceneObjects.find((entry) => entry.id === objectId && entry.type === 'angle');
        if (!item || !item.definition) {
            return;
        }

        const def = item.definition;
        if (!Array.isArray(def.pointIds) || def.pointIds.length !== 3) {
            return;
        }

        const vertexLabel = this.getPointById(def.pointIds[1])?.label || def.pointIds[1];
        const vectors = this.getVectorsByPointIds(def.pointIds);
        if (!vectors || vectors.some((v) => !v)) {
            return;
        }

        let angleLabelInput = def.text || '';
        while (true) {
            const nextLabel = await this.showPromptModal(`Change label for angle at ${vertexLabel}`, angleLabelInput, { quickSymbols: true });
            if (nextLabel == null) {
                return;
            }

            angleLabelInput = nextLabel.trim();
            if (angleLabelInput) {
                break;
            }

            await this.showAlertModal('Angle label cannot be blank.');
        }

        const color = def.color;
        const angleGroup = this.createAngleMarker(vectors[0], vectors[1], vectors[2], angleLabelInput, color);
        item.name = `Angle ${angleLabelInput}`;
        item.subtitle = `Angle at ${vertexLabel}`;
        item.definition = { ...def, text: angleLabelInput };
        this.scene.remove(item.object3D);
        this.disposeObject3D(item.object3D);
        angleGroup.userData.sceneObjectId = item.id;
        item.object3D = angleGroup;
        item.object3D.visible = item.visible;
        this.scene.add(item.object3D);
        this.focusObjectSectionForType(item.type, item.definition);
        this.renderObjectsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    getAngleDisplayHtml(item) {
        const def = item?.definition;
        if (!def || def.kind !== 'angle' || !Array.isArray(def.pointIds) || def.pointIds.length !== 3) {
            return item.name;
        }

        const [aLabel, bLabel, cLabel] = def.pointIds.map((id) => this.getPointById(id)?.label || id);
        return `Angle ${aLabel}${bLabel}\u0302${cLabel}`;
    }

    getSceneObjectDisplayName(item) {
        const definition = item?.definition;
        if (!definition || !Array.isArray(definition.pointIds) || definition.pointIds.length === 0) {
            return item.name;
        }

        if (definition.kind === 'segment' && definition.pointIds.length === 2) {
            return `Segment ${this.formatPointSequence(definition.pointIds)}`;
        }

        if (definition.kind === 'triangle' && definition.pointIds.length === 3) {
            return `Triangle ${this.formatPointSequence(definition.pointIds)}`;
        }

        if (definition.kind === 'angle' && definition.pointIds.length === 3) {
            return `Angle ${this.formatPointSequence(definition.pointIds)}`;
        }

        if (definition.kind === 'plane' && definition.pointIds.length === 4) {
            return `Quadrilateral ${this.formatPointSequence(definition.pointIds)}`;
        }

        if ((definition.kind === 'edge-label' || definition.kind === 'length-label') && definition.pointIds.length === 2) {
            return `Label ${this.formatPointSequence(definition.pointIds)}`;
        }

        if (definition.kind === 'point-label' && definition.pointIds.length === 1) {
            return `Label ${this.formatPointSequence(definition.pointIds)}`;
        }

        return item.name;
    }

    toggleObjectVisibility(objectId) {
        const item = this.sceneObjects.find((entry) => entry.id === objectId);
        if (!item) return;

        item.visible = !item.visible;
        item.object3D.visible = item.visible;
        this.renderObjectsList();
        this.syncEdgeLabelVisibility();
    }

    isEdgePairVisible(pointIds) {
        if (this.hasPrimitiveEdgeBetween(pointIds)) {
            return true;
        }
        if (this.hasSharedDerivedEdgeParent(pointIds)) {
            return true;
        }
        const normalized = this.normalizePointPairIds(pointIds);
        if (!normalized) return false;
        const matchesPair = (pairCandidate) => {
            const pair = this.normalizePointPairIds(pairCandidate);
            return !!pair && pair[0] === normalized[0] && pair[1] === normalized[1];
        };
        return this.sceneObjects.some((entry) => {
            const def = entry.definition;
            if (!def || !Array.isArray(def.pointIds)) return false;
            if (def.kind === 'segment' && def.pointIds.length === 2) {
                if (def.hidden === true) {
                    return matchesPair(def.pointIds);
                }
                if (!entry.visible) return false;
                return matchesPair(def.pointIds);
            }
            if (!entry.visible || def.hidden) return false;
            if (def.kind === 'triangle' && def.pointIds.length === 3) {
                const ids = def.pointIds;
                return matchesPair([ids[0], ids[1]]) || matchesPair([ids[1], ids[2]]) || matchesPair([ids[2], ids[0]]);
            }
            if (def.kind === 'plane' && def.pointIds.length === 4) {
                const ids = def.pointIds;
                return matchesPair([ids[0], ids[1]]) || matchesPair([ids[1], ids[2]]) || matchesPair([ids[2], ids[3]]) || matchesPair([ids[3], ids[0]]);
            }
            return false;
        });
    }

    syncEdgeLabelVisibility() {
        const labelsOn = this.labelMode !== 'off';
        for (const obj of this.sceneObjects) {
            if (obj.type !== 'label') continue;
            const def = obj.definition;
            if (!def || (def.kind !== 'edge-label' && def.kind !== 'length-label')) continue;
            if (!Array.isArray(def.pointIds) || def.pointIds.length !== 2) continue;
            obj.object3D.visible = labelsOn && obj.visible && this.isEdgePairVisible(def.pointIds);
        }
    }

    syncAllLabelVisibility() {
        const labelsOn = this.labelMode !== 'off';
        this.syncEdgeLabelVisibility();
        for (const sprite of (this.pointSprites || [])) {
            sprite.visible = labelsOn;
        }
        for (const obj of this.sceneObjects) {
            if (obj.type !== 'angle') continue;
            const group = obj.object3D;
            if (group && group.children && group.children.length >= 2) {
                group.children[1].visible = labelsOn;
            }
        }
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
        if (this.activeTriangleExtraction?.objectId === objectId) {
            this.closeTriangleExtraction({ force: true });
        }
        if (item.definition?.kind === 'segment' && Array.isArray(item.definition.pointIds) && item.definition.pointIds.length === 2) {
            this.removeEdgeLabelsForPointPair(item.definition.pointIds, { onlyIfDisconnected: true });
            this.removeMidpointPointsForPair(item.definition.pointIds, { onlyIfDisconnected: true });
            this.removeRatioPointsForPair(item.definition.pointIds, { onlyIfDisconnected: true });
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
        // Internal method: clear the 3D scene representation only (not the data)
        if (this.isTriangleExtractionOpen()) {
            this.closeTriangleExtraction({ force: true });
        }

        this.sceneObjects.forEach((item) => {
            this.scene.remove(item.object3D);
            this.disposeObject3D(item.object3D);
        });
        this.sceneObjects = [];
        this.selectedPoints = [];

        this.renderObjectsList();
        this.buildPointMarkers();
        this.renderPointsList();
        this.renderSelectionSummary();
        this.renderActions();
    }

    async clearAllObjects() {
        // User-initiated: clear everything with confirmation
        const confirmed = await this.showConfirmModal(
            'Delete all objects? This cannot be undone.',
            { confirmText: 'Delete', cancelText: 'Cancel' }
        );
        if (!confirmed) {
            return;
        }

        if (this.isTriangleExtractionOpen()) {
            this.closeTriangleExtraction({ force: true });
        }

        // Remove all composite slots properly
        while (this.compositeSlots.length > 0) {
            this.removeSlot(this.compositeSlots[0].id);
        }

        // Clear any remaining scene objects and selections
        this.sceneObjects.forEach((item) => {
            this.scene.remove(item.object3D);
            this.disposeObject3D(item.object3D);
        });
        this.sceneObjects = [];
        this.selectedPoints = [];
        this.pointDefinitions = [];

        // Rebuild UI to show clean state
        this.renderObjectsList();
        this.renderCompositeCards();
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

    _fitCameraParams(object3D, padding = 1.38, preferredDirection = null) {
        if (!object3D) return null;
        const bounds = new THREE.Box3().setFromObject(object3D);
        if (bounds.isEmpty()) return null;
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
        if (viewDir.lengthSq() < 1e-8) viewDir.set(1, 0.72, 0.94);
        viewDir.normalize();
        return { cameraPos: center.clone().add(viewDir.multiplyScalar(distance)), target: center };
    }

    fitCameraToObject(object3D, padding = 1.38, preferredDirection = null) {
        const params = this._fitCameraParams(object3D, padding, preferredDirection);
        if (!params) {
            this.fitCameraToPrimitive(6);
            return;
        }
        this.controls.target.copy(params.target);
        this.camera.position.copy(params.cameraPos);
        this.controls.update();
    }

    resetView() {
        const params = this._fitCameraParams(this.compositeGroup, 1.38, new THREE.Vector3(1, 0.72, 0.94));
        if (!params) {
            this.fitCameraToPrimitive(6);
            return;
        }
        this._cameraAnim = {
            fromPos: this.camera.position.clone(),
            fromTarget: this.controls.target.clone(),
            toPos: params.cameraPos,
            toTarget: params.target,
            startTime: performance.now(),
            duration: 600
        };
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
        if (this._cameraAnim) {
            const { fromPos, fromTarget, toPos, toTarget, startTime, duration } = this._cameraAnim;
            const raw = Math.min(1, (performance.now() - startTime) / duration);
            const t = 1 - Math.pow(1 - raw, 3); // ease-out cubic
            this.camera.position.lerpVectors(fromPos, toPos, t);
            this.controls.target.lerpVectors(fromTarget, toTarget, t);
            if (raw >= 1) this._cameraAnim = null;
        }
        this.controls.update();
        this.updateIntrinsicRightAngleMarkerVisibility();
        this.applyKeyboardCameraInput();
        this.updateEdgeLabelRotations();
        this.renderer.render(this.scene, this.camera);
    }

    applyKeyboardCameraInput() {
        if (!this._keysHeld || this._keysHeld.size === 0) return;
        if (this.triangleExtractOverlay?.classList.contains('show')) return;

        const rotateSpeed = 0.022;
        const zoomSpeed = 0.04;
        const offset = new THREE.Vector3().subVectors(this.camera.position, this.controls.target);
        const spherical = new THREE.Spherical().setFromVector3(offset);

        if (this._keysHeld.has('ArrowLeft'))  spherical.theta -= rotateSpeed;
        if (this._keysHeld.has('ArrowRight')) spherical.theta += rotateSpeed;
        if (this._keysHeld.has('ArrowUp'))    spherical.phi   = Math.max(0.05, spherical.phi - rotateSpeed);
        if (this._keysHeld.has('ArrowDown'))  spherical.phi   = Math.min(Math.PI - 0.05, spherical.phi + rotateSpeed);

        if (this._keysHeld.has('Equal') || this._keysHeld.has('NumpadAdd')) {
            spherical.radius = Math.max(this.controls.minDistance, spherical.radius * (1 - zoomSpeed));
        }
        if (this._keysHeld.has('Minus') || this._keysHeld.has('NumpadSubtract')) {
            spherical.radius = Math.min(this.controls.maxDistance, spherical.radius * (1 + zoomSpeed));
        }

        offset.setFromSpherical(spherical);
        this.camera.position.copy(this.controls.target).add(offset);
        this.controls.update();
    }

    updateEdgeLabelRotations() {
        const aspect = this.canvas.clientWidth / this.canvas.clientHeight;
        const projA = new THREE.Vector3();
        const projB = new THREE.Vector3();
        for (const obj of this.sceneObjects) {
            const def = obj.definition;
            if (!def || (def.kind !== 'edge-label' && def.kind !== 'length-label')) continue;
            const sprite = obj.object3D;
            if (!(sprite instanceof THREE.Sprite)) continue;
            const vectors = this.getVectorsByPointIds(def.pointIds || []);
            if (!vectors || vectors.length !== 2) continue;
            projA.copy(vectors[0]).project(this.camera);
            projB.copy(vectors[1]).project(this.camera);
            const dx = (projB.x - projA.x) * aspect;
            const dy = projB.y - projA.y;
            if (dx * dx + dy * dy < 1e-6) continue;
            let angle = Math.atan2(dy, dx);
            if (angle > Math.PI / 2) angle -= Math.PI;
            if (angle < -Math.PI / 2) angle += Math.PI;
            sprite.material.rotation = angle;
        }
    }

    cleanup() {
        window.removeEventListener('resize', this.handleWindowResize);
        window.removeEventListener('keydown', this._handleKeyDown);
        window.removeEventListener('keyup', this._handleKeyUp);
        this.canvas.removeEventListener('pointerdown', this.handleCanvasPointerDown);
        if (this.crashReportRefreshBtn && this._handleCrashReportRefreshClick) {
            this.crashReportRefreshBtn.removeEventListener('click', this._handleCrashReportRefreshClick);
        }
        if (this.crashReportCopyBtn && this._handleCrashReportCopyClick) {
            this.crashReportCopyBtn.removeEventListener('click', this._handleCrashReportCopyClick);
        }
        if (this.crashReportCloseBtn && this._handleCrashReportCloseClick) {
            this.crashReportCloseBtn.removeEventListener('click', this._handleCrashReportCloseClick);
        }
        if (this.crashReportOverlay && this._handleCrashReportOverlayClick) {
            this.crashReportOverlay.removeEventListener('click', this._handleCrashReportOverlayClick);
        }
        this.closeCrashReport();
        this.teardownCrashDiagnostics();
        this.closeTriangleExtraction({ force: true });
        this.clearObjects();
        this.clearPrimitive();
        window.cancelAnimationFrame(this.animationFrameId);
        this.controls.dispose();
        this.renderer.dispose();
    }
}