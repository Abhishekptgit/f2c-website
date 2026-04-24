
// Gravity Physics Simulation for F2C Website (DOM-Sync Version)
const GravitySim = {
    engine: null,
    world: null,
    runner: null,
    bodies: [],
    isActive: false,

    init() {
        if (typeof Matter === 'undefined') {
            console.error('Matter.js not loaded');
            return;
        }
        const { Engine, Runner, Bodies, Composite } = Matter;

        // Create Engine
        this.engine = Engine.create();
        this.world = this.engine.world;

        // Ground and walls (invisible boundaries)
        const width = window.innerWidth;
        const height = window.innerHeight;
        
        const ground = Bodies.rectangle(width / 2, height + 50, width, 100, { isStatic: true });
        const leftWall = Bodies.rectangle(-50, height / 2, 100, height, { isStatic: true });
        const rightWall = Bodies.rectangle(width + 50, height / 2, 100, height, { isStatic: true });
        
        Composite.add(this.world, [ground, leftWall, rightWall]);

        // Runner
        this.runner = Runner.create();
        Runner.run(this.runner, this.engine);

        // Sync Loop
        Matter.Events.on(this.engine, 'afterUpdate', () => this.syncDOM());
    },

    activate() {
        if (this.isActive) return;
        this.isActive = true;

        const { Bodies, Composite } = Matter;
        const targets = document.querySelectorAll('.hero-title, .hero-subtitle, .btn, .service-card, .trust-item, .logo, .nav-links li');
        
        targets.forEach((el, index) => {
            const rect = el.getBoundingClientRect();
            
            // Create a physics body for each element
            const body = Bodies.rectangle(
                rect.left + rect.width / 2,
                rect.top + rect.height / 2,
                rect.width,
                rect.height,
                {
                    restitution: 0.5,
                    friction: 0.1,
                    chamfer: { radius: 5 } // Rounded corners for smoother sliding
                }
            );

            // Save reference to the DOM element
            body.element = el;
            
            // Prepare the DOM element for absolute positioning
            el.style.position = 'fixed';
            el.style.top = '0';
            el.style.left = '0';
            el.style.width = rect.width + 'px';
            el.style.height = rect.height + 'px';
            el.style.margin = '0';
            el.style.zIndex = '10000';
            el.style.pointerEvents = 'auto';

            Composite.add(this.world, body);
            this.bodies.push(body);
        });

        // Hide the pinned hero background logic if it interferes
        const heroPinned = document.querySelector('.hero-pinned');
        if (heroPinned) heroPinned.style.overflow = 'visible';
    },

    syncDOM() {
        this.bodies.forEach(body => {
            if (body.element) {
                const { x, y } = body.position;
                const angle = body.angle;
                
                // Update DOM position and rotation
                body.element.style.transform = `translate(${x - body.element.offsetWidth / 2}px, ${y - body.element.offsetHeight / 2}px) rotate(${angle}rad)`;
            }
        });
    }
};

window.addEventListener('load', () => {
    GravitySim.init();
});
