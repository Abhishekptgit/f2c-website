
document.addEventListener('DOMContentLoaded', () => {
    // --- Counter Animation ---
    const counters = document.querySelectorAll('.counter');
    const speed = 200;

    const startCounters = () => {
        counters.forEach(counter => {
            const updateCount = () => {
                const target = +counter.getAttribute('data-target');
                const count = +counter.innerText;
                const inc = target / speed;

                if (count < target) {
                    counter.innerText = Math.ceil(count + inc);
                    setTimeout(updateCount, 1);
                } else {
                    counter.innerText = target;
                }
            };
            updateCount();
        });
    };

    // --- GSAP Animations (Snappy Fast Version) ---
    gsap.registerPlugin(ScrollTrigger);

    const images = document.querySelectorAll('.hero-img');
    const texts = document.querySelectorAll('.hero-text-block');

    if (images.length > 0) {
        const heroTl = gsap.timeline({
            scrollTrigger: {
                trigger: ".hero-pinned",
                start: "top top",
                end: "+=5000",
                pin: true,
                scrub: 0.5, // FAST AND SNAPPY
                anticipatePin: 1
            }
        });

        // Set initial states
        gsap.set(images, { opacity: 0 });
        gsap.set(images[0], { opacity: 1 });
        gsap.set(texts, { opacity: 0, y: 50 });
        gsap.set(texts[0], { opacity: 1, y: 0 });
        
        // Simple and fast transitions
        for(let i = 0; i < images.length - 1; i++) {
            heroTl.to(images[i+1], { opacity: 1, duration: 1 }, i)
                  .to(images[i], { opacity: 0, duration: 1 }, i)
                  .to(texts[i], { opacity: 0, y: -50, duration: 0.5 }, i)
                  .to(texts[i+1], { opacity: 1, y: 0, duration: 0.5 }, i + 0.5);
        }
    }

    // Generic reveal animations
    const revealElements = document.querySelectorAll('.gs_reveal, .gs_reveal_up, .gs_reveal_left, .gs_reveal_right');
    
    revealElements.forEach(elem => {
        gsap.to(elem, {
            scrollTrigger: {
                trigger: elem,
                start: "top 90%",
                toggleActions: "play none none none",
                onEnter: () => {
                    if (elem.querySelector('.counter')) {
                        startCounters();
                    }
                }
            },
            opacity: 1,
            y: 0,
            duration: 1.2,
            ease: "power2.out"
        });
    });
});
