/* ============================================================
   OpenLautrec — Page contact
   Validation accessible, envoi via le client mail (mailto),
   et accordéon FAQ.
   ============================================================ */
(function () {
    'use strict';

    const CONTACT_EMAIL = 'kasperweis23@gmail.com';

    const SUBJECT_LABELS = {
        support: 'Support technique',
        bug: 'Signaler un bug',
        feature: 'Suggestion de fonctionnalité',
        education: 'Question éducative',
        partnership: 'Partenariat',
        other: 'Autre'
    };

    /* ---------- Form ---------- */
    function setupForm() {
        const form = document.getElementById('contactForm');
        if (!form) return;
        const success = document.getElementById('successMessage');

        const showError = (field, message) => {
            const error = form.querySelector(`[data-error-for="${field.id}"]`);
            if (error) error.textContent = message;
            field.setAttribute('aria-invalid', message ? 'true' : 'false');
        };

        // Live-clear errors as the user fixes fields.
        form.querySelectorAll('input, select, textarea').forEach((field) => {
            field.addEventListener('input', () => {
                if (field.checkValidity()) showError(field, '');
            });
        });

        form.addEventListener('submit', (e) => {
            e.preventDefault();

            let firstInvalid = null;
            form.querySelectorAll('input, select, textarea').forEach((field) => {
                if (!field.checkValidity()) {
                    showError(field, field.validationMessage);
                    if (!firstInvalid) firstInvalid = field;
                } else {
                    showError(field, '');
                }
            });

            if (firstInvalid) {
                firstInvalid.focus();
                return;
            }

            const name = form.elements.name.value.trim();
            const email = form.elements.email.value.trim();
            const subjectKey = form.elements.subject.value;
            const message = form.elements.message.value.trim();
            const subjectLabel = SUBJECT_LABELS[subjectKey] || 'Contact';

            const body =
                `Nom : ${name}\n` +
                `Email : ${email}\n` +
                `Sujet : ${subjectLabel}\n\n` +
                `${message}\n`;

            const mailto =
                `mailto:${CONTACT_EMAIL}` +
                `?subject=${encodeURIComponent('[OpenLautrec] ' + subjectLabel)}` +
                `&body=${encodeURIComponent(body)}`;

            window.location.href = mailto;

            if (success) {
                success.classList.add('show');
                success.setAttribute('tabindex', '-1');
                success.focus();
            }
            form.reset();

            setTimeout(() => success && success.classList.remove('show'), 8000);
        });
    }

    /* ---------- FAQ accordion ---------- */
    function setupFaq() {
        document.querySelectorAll('.faq-question').forEach((question) => {
            question.addEventListener('click', () => {
                const expanded = question.getAttribute('aria-expanded') === 'true';
                question.setAttribute('aria-expanded', String(!expanded));
                const answer = document.getElementById(question.getAttribute('aria-controls'));
                if (answer) answer.classList.toggle('open', !expanded);
            });
        });
    }

    document.addEventListener('DOMContentLoaded', () => {
        setupForm();
        setupFaq();
    });
})();
