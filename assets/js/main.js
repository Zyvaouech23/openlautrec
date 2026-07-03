/* ============================================================
   OpenLautrec — Scripts partagés (accueil)
   Routage par hash, menu mobile, menu de téléchargement,
   et animations au défilement.
   ============================================================ */
(function () {
    'use strict';

    const VALID_PAGES = ['home', 'features', 'download'];

    /* ---------- Scroll reveal ---------- */
    const reveal = (root = document) => {
        root.querySelectorAll('.fade-in').forEach((el) => el.classList.add('visible'));
    };

    const observer = new IntersectionObserver(
        (entries) => {
            entries.forEach((entry) => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    observer.unobserve(entry.target);
                }
            });
        },
        { threshold: 0.1, rootMargin: '0px 0px -100px 0px' }
    );

    /* ---------- SPA routing (hash based) ---------- */
    function showPage(pageName, updateHash = true) {
        if (!VALID_PAGES.includes(pageName)) {
            pageName = 'home';
        }

        document.querySelectorAll('.page-section').forEach((section) => {
            section.classList.remove('active');
        });

        const target = document.getElementById(pageName + '-page');
        if (target) {
            target.classList.add('active');
            reveal(target);
        }

        // Reflect active state in the navigation for accessibility.
        document.querySelectorAll('[data-route]').forEach((link) => {
            if (link.dataset.route === pageName) {
                link.setAttribute('aria-current', 'page');
            } else {
                link.removeAttribute('aria-current');
            }
        });

        if (updateHash) {
            history.replaceState(null, '', pageName === 'home' ? '#' : '#' + pageName);
        }

        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // Expose for inline handlers / external links if ever needed.
    window.showPage = showPage;

    function currentPageFromHash() {
        return (location.hash || '#home').replace('#', '') || 'home';
    }

    /* ---------- Mobile navigation ---------- */
    function setupMobileNav() {
        const toggle = document.querySelector('.nav-toggle');
        const links = document.getElementById('nav-links');
        if (!toggle || !links) return;

        const closeMenu = () => {
            links.classList.remove('open');
            toggle.setAttribute('aria-expanded', 'false');
        };

        toggle.addEventListener('click', () => {
            const open = links.classList.toggle('open');
            toggle.setAttribute('aria-expanded', String(open));
        });

        links.addEventListener('click', (e) => {
            if (e.target.closest('a')) closeMenu();
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') closeMenu();
        });
    }

    /* ---------- Download split menu ---------- */
    function setupDownloadMenu() {
        const dropdown = document.querySelector('.download-dropdown');
        if (!dropdown) return;
        const btn = dropdown.querySelector('.download-arrow');
        const menu = dropdown.querySelector('.download-menu');
        if (!btn || !menu) return;

        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const open = menu.classList.toggle('open');
            btn.setAttribute('aria-expanded', String(open));
        });

        document.addEventListener('click', (e) => {
            if (!e.target.closest('.download-dropdown')) {
                menu.classList.remove('open');
                btn.setAttribute('aria-expanded', 'false');
            }
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                menu.classList.remove('open');
                btn.setAttribute('aria-expanded', 'false');
            }
        });
    }

    /* ---------- Routing wiring ---------- */
    function setupRouting() {
        document.querySelectorAll('[data-route]').forEach((link) => {
            link.addEventListener('click', (e) => {
                e.preventDefault();
                showPage(link.dataset.route);
            });
        });
        window.addEventListener('hashchange', () => {
            showPage(currentPageFromHash(), false);
        });
    }

    /* ---------- Logo d'arrière-plan pivotant + barre de progression ---------- */
    function setupScrollEffects() {
        const bgLogo = document.querySelector('.bg-logo img');
        const progress = document.querySelector('.scroll-progress');
        const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
        let ticking = false;

        const update = () => {
            ticking = false;
            const scrollY = window.scrollY;
            const max = document.documentElement.scrollHeight - window.innerHeight;

            if (progress) {
                progress.style.width = (max > 0 ? (scrollY / max) * 100 : 0) + '%';
            }

            if (bgLogo && !reduceMotion.matches) {
                const angle = scrollY * 0.12;
                const scale = 1 + Math.min(scrollY / 4000, 0.15);
                bgLogo.style.transform = 'rotate(' + angle + 'deg) scale(' + scale + ')';
            }
        };

        window.addEventListener('scroll', () => {
            if (!ticking) {
                ticking = true;
                requestAnimationFrame(update);
            }
        }, { passive: true });

        update();
    }

    /* ---------- Visite guidée (scrollytelling) ---------- */
    function setupTour() {
        const highlight = document.querySelector('.tour-highlight');
        const steps = document.querySelectorAll('.tour-step');
        if (!highlight || !steps.length) return;

        const activate = (step) => {
            steps.forEach((s) => s.classList.remove('active'));
            step.classList.add('active');
            highlight.classList.add('on');
            highlight.style.left = step.dataset.x + '%';
            highlight.style.top = step.dataset.y + '%';
            highlight.style.width = step.dataset.w + '%';
            highlight.style.height = step.dataset.h + '%';
        };

        const tourObserver = new IntersectionObserver(
            (entries) => {
                entries.forEach((entry) => {
                    if (entry.isIntersecting) activate(entry.target);
                });
            },
            { rootMargin: '-45% 0px -45% 0px', threshold: 0 }
        );

        steps.forEach((step) => tourObserver.observe(step));
    }

    /* ---------- Init ---------- */
    document.addEventListener('DOMContentLoaded', () => {
        document.querySelectorAll('.fade-in').forEach((el) => observer.observe(el));
        setupMobileNav();
        setupDownloadMenu();
        setupRouting();
        setupScrollEffects();
        setupTour();

        // Honour an incoming hash (e.g. from contact.html#features).
        if (document.querySelector('.page-section')) {
            showPage(currentPageFromHash(), false);
        }
    });
})();
