/**
 * 上野高校 SSH コズミックウェブサイト — main.js
 * パーティクルシステム・パララックス・インタラクション全制御
 */

// ===== GLOBAL STATE =====
const state = {
    mouse: { x: window.innerWidth / 2, y: window.innerHeight / 2, prevX: 0, prevY: 0, speed: 0 },
    width: window.innerWidth,
    height: window.innerHeight,
    time: 0,
    grioVariant: 0 // 0 = chemist, 1 = scouter
};

// ===== INIT =====
document.addEventListener('DOMContentLoaded', () => {
    initParallax();
    initSakuraCanvas();
    initParticleCanvas();
    initDataStreamCanvas();
    initPulseBeat();
    initGrioSwap();
    requestAnimationFrame(mainLoop);
});

// ===== MOUSE TRACKING & PARALLAX =====
function initParallax() {
    document.addEventListener('mousemove', (e) => {
        const dx = e.clientX - state.mouse.x;
        const dy = e.clientY - state.mouse.y;
        state.mouse.prevX = state.mouse.x;
        state.mouse.prevY = state.mouse.y;
        state.mouse.x = e.clientX;
        state.mouse.y = e.clientY;
        state.mouse.speed = Math.sqrt(dx * dx + dy * dy);

        // Update custom cursor position
        document.body.style.setProperty('--mouse-x', e.clientX + 'px');
        document.body.style.setProperty('--mouse-y', e.clientY + 'px');

        // Parallax : each layer moves at different intensity
        const cx = (e.clientX / state.width - 0.5);
        const cy = (e.clientY / state.height - 0.5);

        const layerFar = document.getElementById('layerFar');
        const layerMid = document.getElementById('layerMid');
        const layerNear = document.getElementById('layerNear');

        if (layerFar) layerFar.style.transform = `translateZ(-300px) scale(1.25) translate(${cx * 8}px, ${cy * 8}px)`;
        if (layerMid) layerMid.style.transform = `translateZ(0px) translate(${cx * 20}px, ${cy * 20}px)`;
        if (layerNear) layerNear.style.transform = `translateZ(200px) scale(0.83) translate(${cx * 40}px, ${cy * 40}px)`;
    });

    window.addEventListener('resize', () => {
        state.width = window.innerWidth;
        state.height = window.innerHeight;
        resizeCanvases();
    });
}

function resizeCanvases() {
    ['sakuraCanvas', 'particleCanvas', 'dataStreamCanvas'].forEach(id => {
        const cvs = document.getElementById(id);
        if (cvs) {
            cvs.width = state.width;
            cvs.height = state.height;
        }
    });
}

// ===== NEON SAKURA PARTICLES (Layer C) =====
const sakuraParticles = [];
const SAKURA_COUNT = 60;

function initSakuraCanvas() {
    const cvs = document.getElementById('sakuraCanvas');
    if (!cvs) return;
    cvs.width = state.width;
    cvs.height = state.height;

    for (let i = 0; i < SAKURA_COUNT; i++) {
        sakuraParticles.push(createSakura());
    }
}

function createSakura() {
    const isBlue = Math.random() > 0.6;
    return {
        x: Math.random() * state.width,
        y: Math.random() * state.height,
        size: 4 + Math.random() * 10,
        speedX: (Math.random() - 0.5) * 0.5,
        speedY: 0.2 + Math.random() * 0.6,
        rotation: Math.random() * Math.PI * 2,
        rotationSpeed: (Math.random() - 0.5) * 0.03,
        opacity: 0.3 + Math.random() * 0.5,
        color: isBlue
            ? `hsla(200, 100%, 70%, VAR_OPACITY)`
            : `hsla(330, 100%, 70%, VAR_OPACITY)`,
        wobble: Math.random() * Math.PI * 2,
        wobbleSpeed: 0.01 + Math.random() * 0.02
    };
}

function drawSakura(ctx) {
    sakuraParticles.forEach(p => {
        p.x += p.speedX + Math.sin(p.wobble) * 0.3;
        p.y += p.speedY;
        p.rotation += p.rotationSpeed;
        p.wobble += p.wobbleSpeed;

        // Wrap around
        if (p.y > state.height + 20) { p.y = -20; p.x = Math.random() * state.width; }
        if (p.x > state.width + 20) p.x = -20;
        if (p.x < -20) p.x = state.width + 20;

        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate(p.rotation);
        ctx.globalAlpha = p.opacity;

        // Draw petal shape
        const col = p.color.replace('VAR_OPACITY', p.opacity.toString());
        ctx.fillStyle = col;
        ctx.shadowColor = col;
        ctx.shadowBlur = 12;

        ctx.beginPath();
        ctx.ellipse(0, 0, p.size, p.size * 0.5, 0, 0, Math.PI * 2);
        ctx.fill();
        // second petal
        ctx.beginPath();
        ctx.ellipse(0, 0, p.size * 0.5, p.size, 0, 0, Math.PI * 2);
        ctx.fill();

        ctx.restore();
    });
}

// ===== MOUSE TRAIL & SSH SCIENCE PARTICLES (Layer A) =====
const trailParticles = [];
const scienceParticles = [];
const SCIENCE_COUNT = 45;
const SCIENCE_FORMULAS = [
    'E=mc²', 'F=ma', 'ΔG=-nFE', 'PV=nRT', 'λ=h/p',
    '∇×E=-∂B/∂t', 'DNA', 'ATP', 'H₂O', 'CO₂',
    'ψ(x,t)', '∫∫∫', 'Σ', '∞', 'π',
    'd/dx', '∂²ψ/∂x²', 'ℏ', 'Ω', 'μ',
    'C₆H₁₂O₆', 'NaCl', 'Fe₂O₃', 'CH₄'
];

function initParticleCanvas() {
    const cvs = document.getElementById('particleCanvas');
    if (!cvs) return;
    cvs.width = state.width;
    cvs.height = state.height;

    for (let i = 0; i < SCIENCE_COUNT; i++) {
        scienceParticles.push(createScienceParticle());
    }
}

function createScienceParticle() {
    const types = ['formula', 'spark', 'helix'];
    const type = types[Math.floor(Math.random() * types.length)];
    return {
        x: Math.random() * state.width,
        y: Math.random() * state.height,
        baseX: 0,
        baseY: 0,
        vx: (Math.random() - 0.5) * 0.5,
        vy: (Math.random() - 0.5) * 0.5,
        type: type,
        formula: SCIENCE_FORMULAS[Math.floor(Math.random() * SCIENCE_FORMULAS.length)],
        size: type === 'formula' ? (10 + Math.random() * 8) : (2 + Math.random() * 4),
        opacity: 0.4 + Math.random() * 0.4,
        color: type === 'spark'
            ? `hsl(${120 + Math.random() * 40}, 100%, 60%)`
            : `hsl(${170 + Math.random() * 50}, 100%, 70%)`,
        phase: Math.random() * Math.PI * 2,
        repelled: false
    };
}

function drawParticles(ctx) {
    const mx = state.mouse.x;
    const my = state.mouse.y;
    const repelRadius = 150;

    scienceParticles.forEach(p => {
        // Mouse repulsion
        const dx = p.x - mx;
        const dy = p.y - my;
        const dist = Math.sqrt(dx * dx + dy * dy);

        if (dist < repelRadius && dist > 0) {
            const force = (repelRadius - dist) / repelRadius * 8;
            p.vx += (dx / dist) * force;
            p.vy += (dy / dist) * force;
            p.repelled = true;
        } else {
            p.repelled = false;
        }

        // Friction
        p.vx *= 0.95;
        p.vy *= 0.95;

        // Drift
        p.vx += (Math.random() - 0.5) * 0.1;
        p.vy += (Math.random() - 0.5) * 0.1;

        p.x += p.vx;
        p.y += p.vy;
        p.phase += 0.02;

        // Wrap
        if (p.x < -50) p.x = state.width + 50;
        if (p.x > state.width + 50) p.x = -50;
        if (p.y < -50) p.y = state.height + 50;
        if (p.y > state.height + 50) p.y = -50;

        ctx.save();
        ctx.globalAlpha = p.opacity + (p.repelled ? 0.3 : 0);

        if (p.type === 'formula') {
            ctx.font = `${p.size}px 'Orbitron', monospace`;
            ctx.fillStyle = p.color;
            ctx.shadowColor = p.color;
            ctx.shadowBlur = p.repelled ? 25 : 10;
            ctx.fillText(p.formula, p.x, p.y);
        } else if (p.type === 'spark') {
            // Green spark
            ctx.fillStyle = p.color;
            ctx.shadowColor = p.color;
            ctx.shadowBlur = p.repelled ? 20 : 8;
            ctx.beginPath();
            ctx.arc(p.x, p.y, p.size * (p.repelled ? 2 : 1), 0, Math.PI * 2);
            ctx.fill();
        } else if (p.type === 'helix') {
            // DNA helix strand dots
            const helixY1 = p.y + Math.sin(p.phase) * 15;
            const helixY2 = p.y + Math.sin(p.phase + Math.PI) * 15;
            ctx.fillStyle = '#00ff88';
            ctx.shadowColor = '#00ff88';
            ctx.shadowBlur = 8;
            ctx.beginPath();
            ctx.arc(p.x - 5, helixY1, 3, 0, Math.PI * 2);
            ctx.fill();
            ctx.fillStyle = '#ff00ff';
            ctx.shadowColor = '#ff00ff';
            ctx.beginPath();
            ctx.arc(p.x + 5, helixY2, 3, 0, Math.PI * 2);
            ctx.fill();
            // connecting line
            ctx.strokeStyle = 'rgba(0, 240, 255, 0.3)';
            ctx.shadowBlur = 4;
            ctx.lineWidth = 1;
            ctx.beginPath();
            ctx.moveTo(p.x - 5, helixY1);
            ctx.lineTo(p.x + 5, helixY2);
            ctx.stroke();
        }

        ctx.restore();
    });

    // Mouse trail (digital smoke)
    if (state.mouse.speed > 2) {
        for (let i = 0; i < Math.min(state.mouse.speed * 0.5, 5); i++) {
            trailParticles.push({
                x: mx + (Math.random() - 0.5) * 20,
                y: my + (Math.random() - 0.5) * 20,
                vx: (Math.random() - 0.5) * 3,
                vy: (Math.random() - 0.5) * 3,
                life: 1,
                decay: 0.02 + Math.random() * 0.03,
                size: 2 + Math.random() * 6,
                hue: 300 + Math.random() * 60 // pink-magenta range
            });
        }
    }

    // Draw & update trail
    for (let i = trailParticles.length - 1; i >= 0; i--) {
        const tp = trailParticles[i];
        tp.x += tp.vx;
        tp.y += tp.vy;
        tp.vx *= 0.96;
        tp.vy *= 0.96;
        tp.life -= tp.decay;
        tp.size *= 1.01;

        if (tp.life <= 0) {
            trailParticles.splice(i, 1);
            continue;
        }

        ctx.save();
        ctx.globalAlpha = tp.life * 0.7;
        const col = `hsl(${tp.hue}, 100%, 65%)`;
        ctx.fillStyle = col;
        ctx.shadowColor = col;
        ctx.shadowBlur = 15;
        ctx.beginPath();
        ctx.arc(tp.x, tp.y, tp.size, 0, Math.PI * 2);
        ctx.fill();

        // Spark cross
        if (tp.life > 0.6 && tp.size < 6) {
            ctx.strokeStyle = `hsla(${tp.hue}, 100%, 80%, ${tp.life})`;
            ctx.lineWidth = 1;
            ctx.beginPath();
            ctx.moveTo(tp.x - tp.size * 2, tp.y);
            ctx.lineTo(tp.x + tp.size * 2, tp.y);
            ctx.moveTo(tp.x, tp.y - tp.size * 2);
            ctx.lineTo(tp.x, tp.y + tp.size * 2);
            ctx.stroke();
        }

        ctx.restore();
    }
}

// ===== DATA STREAM (Layer C) =====
function initDataStreamCanvas() {
    const cvs = document.getElementById('dataStreamCanvas');
    if (!cvs) return;
    cvs.width = state.width;
    cvs.height = state.height;
}

const dataStreams = [];
const DATA_STREAM_TARGETS = [
    () => ({ x: state.width * 0.08, y: state.height * 0.15 }),
    () => ({ x: state.width * 0.88, y: state.height * 0.10 }),
    () => ({ x: state.width * 0.15, y: state.height * 0.60 }),
    () => ({ x: state.width * 0.82, y: state.height * 0.55 }),
    () => ({ x: state.width * 0.50, y: state.height * 0.80 }),
    () => ({ x: state.width * 0.25, y: state.height * 0.20 }),
    () => ({ x: state.width * 0.70, y: state.height * 0.35 }),
];

function drawDataStreams(ctx) {
    const cx = state.width / 2;
    const cy = state.height / 2;
    const t = state.time * 0.001;

    DATA_STREAM_TARGETS.forEach((getTarget, i) => {
        const target = getTarget();

        ctx.save();

        // Glitchy line
        const segments = 20;
        ctx.strokeStyle = `hsla(200, 100%, 60%, ${0.15 + Math.sin(t * 2 + i) * 0.1})`;
        ctx.lineWidth = 1;
        ctx.shadowColor = 'rgba(79, 195, 247, 0.5)';
        ctx.shadowBlur = 8;

        ctx.beginPath();
        ctx.moveTo(cx, cy);

        for (let s = 1; s <= segments; s++) {
            const ratio = s / segments;
            const lx = cx + (target.x - cx) * ratio;
            const ly = cy + (target.y - cy) * ratio;
            // Add glitch jitter
            const glitch = Math.sin(t * 10 + i * 3 + s) * 3 * (1 - ratio);
            ctx.lineTo(lx + glitch, ly + glitch);
        }
        ctx.stroke();

        // Data pulse traveling along the line
        const pulsePos = ((t * 0.5 + i * 0.3) % 1);
        const px = cx + (target.x - cx) * pulsePos;
        const py = cy + (target.y - cy) * pulsePos;

        ctx.fillStyle = `hsla(200, 100%, 70%, ${0.6 + Math.sin(t * 5) * 0.3})`;
        ctx.shadowColor = 'rgba(79, 195, 247, 0.8)';
        ctx.shadowBlur = 15;
        ctx.beginPath();
        ctx.arc(px, py, 3, 0, Math.PI * 2);
        ctx.fill();

        ctx.restore();
    });
}

// ===== PULSE BEAT (和太鼓リズム) =====
function initPulseBeat() {
    const garden = document.getElementById('meijiGarden');
    if (garden) {
        garden.classList.add('pulsing');
    }
}

// ===== GRIO IMAGE SWAP =====
function initGrioSwap() {
    const pet = document.getElementById('grioPet');
    if (!pet) return;

    // Swap grio variants every 10 seconds
    setInterval(() => {
        const imgA = document.getElementById('grioImgA');
        const imgB = document.getElementById('grioImgB');
        if (!imgA || !imgB) return;

        state.grioVariant = 1 - state.grioVariant;
        if (state.grioVariant === 0) {
            imgA.classList.add('grio-current');
            imgB.classList.remove('grio-current');
        } else {
            imgB.classList.add('grio-current');
            imgA.classList.remove('grio-current');
        }
    }, 10000);
}

// ===== MAIN ANIMATION LOOP =====
function mainLoop(timestamp) {
    state.time = timestamp;

    // Sakura canvas
    const sakuraCvs = document.getElementById('sakuraCanvas');
    if (sakuraCvs) {
        const sCtx = sakuraCvs.getContext('2d');
        sCtx.clearRect(0, 0, state.width, state.height);
        drawSakura(sCtx);
    }

    // Particle canvas (science + trail)
    const particleCvs = document.getElementById('particleCanvas');
    if (particleCvs) {
        const pCtx = particleCvs.getContext('2d');
        pCtx.clearRect(0, 0, state.width, state.height);
        drawParticles(pCtx);
    }

    // Data stream canvas
    const dsCvs = document.getElementById('dataStreamCanvas');
    if (dsCvs) {
        const dsCtx = dsCvs.getContext('2d');
        dsCtx.clearRect(0, 0, state.width, state.height);
        drawDataStreams(dsCtx);
    }

    requestAnimationFrame(mainLoop);
}
