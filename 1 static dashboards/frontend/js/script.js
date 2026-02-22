class RentalApplication {
    constructor() {
        this.config = {
            LOCAL_STORAGE_KEY: "choicePropertiesRentalApp",
            AUTO_SAVE_INTERVAL: 30000,
            MAX_FILE_SIZE: 10 * 1024 * 1024
        };
        
        this.state = {
            currentSection: 1,
            isSubmitting: false,
            isOnline: true,
            lastSave: null,
            applicationId: null,
            formData: {}
        };
        
        this.BACKEND_URL = 'https://script.google.com/macros/s/AKfycbz6F0REVb14ocphZTf45XFpm7qPhCQGZuRq7NGsAOo6MOViZrEwThxL7DJrRgN5KVth/exec';
        
        this.initialize();
}

    // ---------- SSN toggle ----------
    setupSSNToggle() {
        const ssnInput = document.getElementById('ssn');
        if (!ssnInput) return;

        const container = ssnInput.parentElement;
        let toggle = container.querySelector('.ssn-toggle');
        
        if (!toggle) {
            toggle = document.createElement('button');
            toggle.type = 'button';
            toggle.className = 'ssn-toggle';
            toggle.id = 'ssnToggle';
            toggle.innerHTML = '<i class="fas fa-eye"></i>';
            container.appendChild(toggle);
        }

        ssnInput.type = 'password';

        toggle.addEventListener('click', () => {
            if (ssnInput.type === 'password') {
                ssnInput.type = 'text';
                toggle.innerHTML = '<i class="fas fa-eye-slash"></i>';
            } else {
                ssnInput.type = 'password';
                toggle.innerHTML = '<i class="fas fa-eye"></i>';
            }
        });

        ssnInput.addEventListener('input', (e) => {
            e.target.value = e.target.value.replace(/\D/g, '').substring(0, 4);
        });
    }

    // ---------- Event listeners ----------
    setupEventListeners() {
        const payBtn = document.getElementById('payNowBtn');
        if (payBtn) {
            payBtn.addEventListener('click', () => {
                alert('Redirecting to secure payment checkout...');
            });
        }
        
        const mainSubmitBtn = document.getElementById('mainSubmitBtn');
        if (mainSubmitBtn) {
            mainSubmitBtn.addEventListener('click', (e) => {
                const certify = document.getElementById('certifyCorrect');
                const authorize = document.getElementById('authorizeVerify');
                const terms = document.getElementById('termsAgree');
                
                if (!certify.checked || !authorize.checked || !terms.checked) {
                    e.preventDefault();
                    alert('Please agree to all legal declarations before submitting.');
                }
            });
        }

        // Navigation
        document.addEventListener('click', (e) => {
            if (e.target.matches('.btn-next') || e.target.closest('.btn-next')) {
                const section = this.getCurrentSection();
                this.nextSection(section);
            }
            if (e.target.matches('.btn-prev') || e.target.closest('.btn-prev')) {
                const section = this.getCurrentSection();
                this.previousSection(section);
            }
        });
        
        // Auto-save on input (debounced)
        document.addEventListener('input', this.debounce(() => {
            this.saveProgress();
        }, 1000));
        
        // Form submission
        const form = document.getElementById('rentalApplication');
        if (form) {
            form.addEventListener('submit', (e) => {
                this.handleFormSubmit(e);
            });
        }
    }

    // ---------- Initialization ----------
    initialize() {
        this.setupEventListeners();
        this.setupOfflineDetection();
        this.setupRealTimeValidation();
        this.setupSSNToggle();
        this.setupFileUploads();
        this.setupConditionalFields();
        this.setupCharacterCounters();
        this.restoreSavedProgress();
        this.setupGeoapify();
        this.setupInputFormatting();
        this.setupLanguageToggle();
        
        console.log('Rental Application Manager Initialized');
    }

    // ---------- Offline detection ----------
    setupOfflineDetection() {
        window.addEventListener('online', () => {
            this.setState({ isOnline: true });
        });
        
        window.addEventListener('offline', () => {
            this.setState({ isOnline: false });
        });
        
        this.setState({ isOnline: navigator.onLine });
    }

    setState(newState) {
        this.state = { ...this.state, ...newState };
        this.updateUIState();
    }

    updateUIState() {
        const offlineIndicator = document.getElementById('offlineIndicator');
        if (offlineIndicator) {
            offlineIndicator.style.display = this.state.isOnline ? 'none' : 'block';
        }
    }

    // ---------- Geoapify ----------
    setupGeoapify() {
        const apiKey = "bea2afb13c904abea5cb2c2693541dcf";
        const fields = ['propertyAddress', 'currentAddress'];
        
        fields.forEach(id => {
            const input = document.getElementById(id);
            if (!input) return;

            const container = document.createElement('div');
            container.style.position = 'relative';
            input.parentNode.insertBefore(container, input);
            container.appendChild(input);

            const dropdown = document.createElement('div');
            dropdown.className = 'autocomplete-dropdown';
            dropdown.style.cssText = 'position: absolute; top: 100%; left: 0; right: 0; background: white; border: 1px solid #ddd; z-index: 1000; display: none; max-height: 200px; overflow-y: auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-radius: 4px;';
            container.appendChild(dropdown);

            input.addEventListener('input', this.debounce(async (e) => {
                const text = e.target.value;
                if (text.length < 3) {
                    dropdown.style.display = 'none';
                    return;
                }

                try {
                    const response = await fetch(`https://api.geoapify.com/v1/geocode/autocomplete?text=${encodeURIComponent(text)}&apiKey=${apiKey}`);
                    const data = await response.json();
                    
                    if (data.features && data.features.length > 0) {
                        dropdown.innerHTML = '';
                        data.features.forEach(feature => {
                            const item = document.createElement('div');
                            item.style.cssText = 'padding: 10px; cursor: pointer; border-bottom: 1px solid #eee; font-size: 14px;';
                            item.textContent = feature.properties.formatted;
                            item.addEventListener('mouseover', () => item.style.background = '#f0f7ff');
                            item.addEventListener('mouseout', () => item.style.background = 'white');
                            item.addEventListener('click', () => {
                                input.value = feature.properties.formatted;
                                dropdown.style.display = 'none';
                                this.saveProgress();
                            });
                            dropdown.appendChild(item);
                        });
                        dropdown.style.display = 'block';
                    } else {
                        dropdown.style.display = 'none';
                    }
                } catch (err) {
                    console.error('Geocoding error:', err);
                }
            }, 300));

            document.addEventListener('click', (e) => {
                if (!container.contains(e.target)) dropdown.style.display = 'none';
            });
        });
    }

    // ---------- Input formatting (phone, SSN) ----------
    setupInputFormatting() {
        const phoneFields = ['phone', 'landlordPhone', 'supervisorPhone', 'ref1Phone', 'ref2Phone', 'emergencyPhone', 'coPhone'];
        phoneFields.forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('input', (e) => {
                    let x = e.target.value.replace(/\D/g, '').match(/(\d{0,3})(\d{0,3})(\d{0,4})/);
                    e.target.value = !x[2] ? x[1] : '(' + x[1] + ') ' + x[2] + (x[3] ? '-' + x[3] : '');
                });
            }
        });

        const ssnEl = document.getElementById('ssn');
        if (ssnEl) {
            ssnEl.addEventListener('input', (e) => {
                let val = e.target.value.replace(/\D/g, '');
                if (val.length > 4) val = val.substring(0, 4);
                e.target.value = val;
                if (val.length === 4) {
                    this.clearError(ssnEl);
                }
            });
            ssnEl.addEventListener('blur', () => {
                this.validateField(ssnEl);
            });
        }

        const coSsnEl = document.getElementById('coSsn');
        if (coSsnEl) {
            coSsnEl.addEventListener('input', (e) => {
                let val = e.target.value.replace(/\D/g, '');
                if (val.length > 4) val = val.substring(0, 4);
                e.target.value = val;
                if (val.length === 4) {
                    this.clearError(coSsnEl);
                }
            });
        }
    }

    // ---------- Real-time validation ----------
    setupRealTimeValidation() {
        const form = document.getElementById('rentalApplication');
        if (!form) return;
        const inputs = form.querySelectorAll('input, select, textarea');
        inputs.forEach(input => {
            input.addEventListener('input', () => this.validateField(input));
            input.addEventListener('change', () => this.validateField(input));
            input.addEventListener('blur', () => this.validateField(input));
        });
    }

    // ---------- Validation logic ----------
    validateField(field) {
        let isValid = true;
        let errorMessage = 'Required';
        
        if (field.id === 'ssn' || field.id === 'coSsn') {
            const ssnVal = field.value.replace(/\D/g, '');
            if (!ssnVal) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'Please enter the last 4 digits of your SSN.' : 'Por favor ingrese los últimos 4 dígitos de su SSN.';
            } else if (ssnVal.length < 4) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'SSN must contain 4 digits.' : 'El SSN debe contener 4 dígitos.';
            } else if (/[^0-9]/.test(field.value)) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'SSN must contain numbers only.' : 'El SSN debe contener solo números.';
            }
        } else if (field.id === 'dob' || field.id === 'coDob') {
            const birthDate = new Date(field.value);
            const today = new Date();
            
            if (!field.value) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'Please enter your date of birth.' : 'Por favor ingrese su fecha de nacimiento.';
            } else if (isNaN(birthDate.getTime())) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'Please enter a valid date of birth (18+ required).' : 'Por favor ingrese una fecha válida (18+ requerido).';
            } else {
                let age = today.getFullYear() - birthDate.getFullYear();
                const m = today.getMonth() - birthDate.getMonth();
                if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
                    age--;
                }
                if (age < 18) {
                    isValid = false;
                    errorMessage = this.state.language === 'en' ? 'Applicants must be at least 18 years old.' : 'Los solicitantes deben tener al menos 18 años.';
                }
            }
        } else if (field.id === 'requestedMoveIn') {
            const moveInDate = new Date(field.value);
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            
            if (!field.value) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'Please select a move-in date.' : 'Por favor seleccione una fecha de mudanza.';
            } else if (moveInDate < today) {
                isValid = false;
                errorMessage = this.state.language === 'en' ? 'Move-in date cannot be in the past.' : 'La fecha de mudanza no puede ser en el pasado.';
            }
        } else if (field.hasAttribute('required')) {
            if (field.type === 'checkbox') {
                isValid = field.checked;
            } else if (!field.value.trim()) {
                isValid = false;
            }
            
            if (!isValid) {
                errorMessage = this.state.language === 'en' ? 'Required' : 'Campo obligatorio';
            }
        }
        
        if (isValid && field.value.trim()) {
            if (field.type === 'email') {
                const email = field.value.trim();
                if (!email.includes('@')) {
                    isValid = false;
                    errorMessage = this.state.language === 'en' ? 'Email must include an @ symbol.' : 'El correo debe incluir un símbolo @.';
                } else {
                    const parts = email.split('@');
                    if (!parts[1] || !parts[1].includes('.')) {
                        isValid = false;
                        errorMessage = this.state.language === 'en' ? 'Add a valid domain (e.g., gmail.com).' : 'Agregue un dominio válido (ej. gmail.com).';
                    } else {
                        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
                        isValid = emailRegex.test(email);
                        if (!isValid) {
                            errorMessage = this.state.language === 'en' ? 'Enter a valid email (example: name@email.com).' : 'Ingrese un correo válido (ejemplo: nombre@email.com).';
                        }
                    }
                }
            } else if (field.type === 'tel') {
                const phoneDigits = field.value.replace(/\D/g, '');
                isValid = phoneDigits.length >= 10;
                if (!isValid) {
                    errorMessage = this.state.language === 'en' ? 'Invalid phone' : 'Teléfono inválido';
                }
            }
        }

        if (isValid) {
            this.clearError(field);
            field.classList.add('is-valid');
            field.classList.remove('is-invalid');
        } else {
            this.showError(field, errorMessage);
            field.classList.add('is-invalid');
            field.classList.remove('is-valid');
            field.classList.add('shake');
            setTimeout(() => field.classList.remove('shake'), 400);
        }
        
        return isValid;
    }

    showError(field, message) {
        field.classList.add('error');
        const errorMsg = field.closest('.form-group')?.querySelector('.error-message');
        if (errorMsg) {
            errorMsg.textContent = message;
            errorMsg.style.display = 'block';
        }
    }

    clearError(field) {
        field.classList.remove('error');
        const errorMsg = field.closest('.form-group')?.querySelector('.error-message');
        if (errorMsg) {
            errorMsg.style.display = 'none';
        }
    }

    // ---------- Section navigation ----------
    getCurrentSection() {
        const activeSection = document.querySelector('.form-section.active');
        return activeSection ? parseInt(activeSection.id.replace('section', '')) : 1;
    }

    nextSection(currentSection) {
        if (!this.validateStep(currentSection)) {
            return;
        }
        
        this.hideSection(currentSection);
        this.showSection(currentSection + 1);
        this.updateProgressBar();
        
        if (currentSection + 1 === 6) {
            this.generateApplicationSummary();
        }
    }

    previousSection(currentSection) {
        if (currentSection > 1) {
            this.hideSection(currentSection);
            this.showSection(currentSection - 1);
            this.updateProgressBar();
        }
    }

    hideSection(sectionNumber) {
        const section = document.getElementById(`section${sectionNumber}`);
        if (section) section.classList.remove('active');
    }

    showSection(sectionNumber) {
        const section = document.getElementById(`section${sectionNumber}`);
        if (section) {
            section.classList.add('active');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    }

    updateProgressBar() {
        const currentSection = this.getCurrentSection();
        const progress = ((currentSection - 1) / 5) * 100;
        const progressFill = document.getElementById('progressFill');
        if (progressFill) {
            progressFill.style.width = `${progress}%`;
        }
        
        const progressContainer = document.querySelector('.progress-container');
        // Get translations for step labels and progress text
        const t = this.getTranslations();
        const stepNames = [
            t.step1Label,
            t.step2Label,
            t.step3Label,
            t.step4Label,
            t.step5Label,
            t.step6Label
        ];
        const progressText = `${t.stepPrefix} ${currentSection} ${t.stepOf} 6: ${stepNames[currentSection-1]}`;
        if (progressContainer) {
            progressContainer.setAttribute('data-progress-text', progressText);
        }
        
        for (let i = 1; i <= 6; i++) {
            const step = document.getElementById(`step${i}`);
            if (step) {
                step.classList.remove('active', 'completed');
                if (i < currentSection) step.classList.add('completed');
                if (i === currentSection) step.classList.add('active');
            }
        }
    }

    // ---------- Step validation ----------
    validateStep(stepNumber) {
        if (stepNumber === 5) {
            const isUnique = this.validatePaymentSelections();
            if (!isUnique) {
                const warning = document.getElementById('paymentDuplicateWarning');
                if (warning) {
                    warning.classList.add('shake');
                    setTimeout(() => warning.classList.remove('shake'), 400);
                }
                return false;
            }
        }

        const section = document.getElementById(`section${stepNumber}`);
        if (!section) return true;
        
        const inputs = section.querySelectorAll('input, select, textarea');
        let isStepValid = true;
        let firstInvalidField = null;
        
        inputs.forEach(input => {
            if (input.hasAttribute('required')) {
                if (!this.validateField(input)) {
                    isStepValid = false;
                    if (!firstInvalidField) firstInvalidField = input;
                }
            }
        });

        if (stepNumber === 1) {
            const hasCoApplicant = document.getElementById('hasCoApplicant');
            const coSection = document.getElementById('coApplicantSection');
            if (hasCoApplicant && hasCoApplicant.checked && coSection && coSection.style.display !== 'none') {
                const coInputs = coSection.querySelectorAll('input, select, textarea');
                coInputs.forEach(input => {
                    if (input.type === 'radio') {
                        const name = input.name;
                        const radios = coSection.querySelectorAll(`input[name="${name}"]`);
                        const checked = Array.from(radios).some(r => r.checked);
                        if (!checked) {
                            this.showError(radios[0], this.state.language === 'en' ? 'Please select a role' : 'Por favor seleccione un rol');
                            radios[0].classList.add('is-invalid');
                            isStepValid = false;
                            if (!firstInvalidField) firstInvalidField = radios[0];
                        } else {
                            radios.forEach(r => this.clearError(r));
                        }
                    } else if (input.type === 'checkbox') {
                        if (input.id === 'coConsent' && !input.checked) {
                            this.showError(input, this.state.language === 'en' ? 'You must authorise verification' : 'Debe autorizar la verificación');
                            input.classList.add('is-invalid');
                            isStepValid = false;
                            if (!firstInvalidField) firstInvalidField = input;
                        } else {
                            this.clearError(input);
                        }
                    } else {
                        if (!input.value.trim()) {
                            this.showError(input, this.state.language === 'en' ? 'Required' : 'Campo obligatorio');
                            input.classList.add('is-invalid');
                            isStepValid = false;
                            if (!firstInvalidField) firstInvalidField = input;
                        } else {
                            if (!this.validateField(input)) {
                                isStepValid = false;
                                if (!firstInvalidField) firstInvalidField = input;
                            }
                        }
                    }
                });
            }
        }

        if (!isStepValid && firstInvalidField) {
            this.scrollToInvalidField(firstInvalidField);
        }
        
        return isStepValid;
    }

    validatePaymentSelections() {
        const s1 = document.getElementById('primaryPayment').value;
        const s2 = document.getElementById('secondaryPayment').value;
        const s3 = document.getElementById('thirdPayment').value;
        const warning = document.getElementById('paymentDuplicateWarning');
        
        let hasDuplicate = false;
        const values = [s1, s2, s3].filter(v => v && v !== 'Other');
        const uniqueValues = new Set(values);
        
        if (values.length !== uniqueValues.size) {
            hasDuplicate = true;
        }

        if (warning) {
            warning.style.display = hasDuplicate ? 'block' : 'none';
        }
        
        return !hasDuplicate;
    }

    scrollToInvalidField(field) {
        const scrollTarget = field.closest('.form-group') || field;
        scrollTarget.scrollIntoView({ behavior: 'smooth', block: 'center' });
        field.classList.add('shake', 'highlight-field');
        setTimeout(() => {
            field.focus();
        }, 600);
        setTimeout(() => {
            field.classList.remove('shake', 'highlight-field');
        }, 2000);
    }

    // ---------- Conditional fields ----------
    setupConditionalFields() {
        const paymentSelectors = ['primaryPayment', 'secondaryPayment', 'thirdPayment'];
        paymentSelectors.forEach(id => {
            const select = document.getElementById(id);
            if (select) {
                select.addEventListener('change', (e) => {
                    const otherContainer = document.getElementById(`${id}OtherContainer`);
                    if (otherContainer) {
                        otherContainer.style.display = e.target.value === 'Other' ? 'block' : 'none';
                    }
                    this.validatePaymentSelections();
                });
            }
        });

        const petsRadio = document.getElementsByName('Has Pets');
        const petGroup = document.getElementById('petDetailsGroup');
        if (petsRadio && petGroup) {
            petsRadio.forEach(r => r.addEventListener('change', (e) => {
                petGroup.style.display = e.target.value === 'Yes' ? 'block' : 'none';
            }));
        }

        const hasCoApplicantCheck = document.getElementById('hasCoApplicant');
        const coApplicantSection = document.getElementById('coApplicantSection');
        if (hasCoApplicantCheck && coApplicantSection) {
            hasCoApplicantCheck.addEventListener('change', (e) => {
                coApplicantSection.style.display = e.target.checked ? 'block' : 'none';
                if (!e.target.checked) {
                    const inputs = coApplicantSection.querySelectorAll('input, select, textarea');
                    inputs.forEach(input => this.clearError(input));
                }
            });
        }

        const vehicleYes = document.getElementById('vehicleYes');
        const vehicleNo = document.getElementById('vehicleNo');
        const vehicleDetails = document.getElementById('vehicleDetailsSection');
        if (vehicleYes && vehicleNo && vehicleDetails) {
            const toggleVehicle = () => {
                vehicleDetails.style.display = vehicleYes.checked ? 'block' : 'none';
            };
            vehicleYes.addEventListener('change', toggleVehicle);
            vehicleNo.addEventListener('change', toggleVehicle);
        }
    }

    setupFileUploads() {}

    setupCharacterCounters() {
        const textareas = document.querySelectorAll('textarea');
        textareas.forEach(textarea => {
            const parent = textarea.parentElement;
            const counter = document.createElement('div');
            counter.className = 'character-count';
            counter.style.fontSize = '11px';
            counter.style.textAlign = 'right';
            counter.style.color = '#7f8c8d';
            parent.appendChild(counter);

            const updateCounter = () => {
                const len = textarea.value.length;
                const max = textarea.getAttribute('maxlength') || 500;
                counter.textContent = `${len}/${max} characters`;
            };

            textarea.addEventListener('input', updateCounter);
            updateCounter();
        });
    }

    restoreSavedProgress() {
        const saved = localStorage.getItem(this.config.LOCAL_STORAGE_KEY);
        if (saved) {
            try {
                const data = JSON.parse(saved);
                Object.keys(data).forEach(key => {
                    if (key === 'SSN' || key === 'Co-Applicant SSN') return;
                    const el = document.getElementById(key);
                    if (el) {
                        if (el.type === 'checkbox') el.checked = data[key];
                        else el.value = data[key];
                    }
                });
                if (data._language) this.state.language = data._language;
            } catch (e) {}
        }
    }

    saveProgress() {
        const data = this.getAllFormData();
        const sensitiveKeys = ['SSN', 'Application ID', 'Co-Applicant SSN'];
        sensitiveKeys.forEach(key => delete data[key]);
        
        data._last_updated = new Date().toISOString();
        data._language = this.state.language || 'en';
        
        localStorage.setItem(this.config.LOCAL_STORAGE_KEY, JSON.stringify(data));
    }

    getAllFormData() {
        const form = document.getElementById('rentalApplication');
        const formData = new FormData(form);
        const data = {};
        formData.forEach((value, key) => {
            data[key] = value;
        });
        return data;
    }

    debounce(func, wait) {
        let timeout;
        return function() {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, arguments), wait);
        };
    }

    // ---------- Language toggle (now uses data-i18n) ----------
    setupLanguageToggle() {
        const translations = {
            en: {
                langText: 'Español',
                logoText: 'Choice Properties',
                tagline: 'Professional Property Management Solutions',
                confidentialStamp: 'CONFIDENTIAL & SECURE',
                trustIndicator: 'Your information is encrypted and protected',
                timeEstimate: 'Estimated time: 15-20 minutes',
                step1Label: 'Property & Applicant',
                step2Label: 'Residency & Occupancy',
                step3Label: 'Employment & Income',
                step4Label: 'Financial & References',
                step5Label: 'Payment Preferences',
                step6Label: 'Review & Submit',
                stepPrefix: 'Step',
                stepOf: 'of',
                processing: 'Processing',
                validating: 'Validating',
                submitting: 'Submitting',
                complete: 'Complete',
                submittingTitle: 'Submitting Your Application',
                submissionMessage: "Please don't close this window. This may take a few moments...",
                successTitle: 'Application Received',
                successText: 'Thank you for choosing Choice Properties',
                appId: 'Your Application ID',
                clickToCopy: 'Copy ID',
                immediateNextSteps: 'Immediate Next Steps',
                paymentRequiredTitle: 'Payment Required Before Review',
                paymentRequiredDesc: 'Our team will contact you shortly at the phone number provided to arrange the $50 application fee.',
                completePaymentTitle: 'Complete Payment',
                completePaymentDesc: 'Your application is not complete until the $50 fee has been paid. We\'ll discuss payment options you\'re familiar with.',
                reviewBeginsTitle: 'Review Begins',
                reviewBeginsDesc: 'Once payment is confirmed, your application enters the formal review process. You can track status online with your ID.',
                importantNote: 'Important:',
                paymentUrgentText: 'Your application is not complete until the $50 fee has been paid. Please keep your phone nearby.',
                yourPreferences: 'Your Preferences (For Follow-up After Payment)',
                contactMethod: 'Contact Method:',
                bestTimes: 'Best Times:',
                paymentPref: 'Payment Preferences:',
                preferenceNote: 'We\'ll use these for non-urgent follow-up after your payment is complete.',
                questions: 'Questions? Call or text',
                helpText: 'we\'re here to help.',
                trackStatus: 'Track My Application',
                newApplication: 'New Application',
                reapplicationPolicyTitle: 'Reapplication Protection',
                reapplicationPolicyText: 'If your application is denied, you may apply for any other available property within 30 days — no new application fee. Your screening results remain valid for 60 days.',
                step1YouSubmit: '1. You Submit',
                step1Desc: 'Fill out your application completely',
                step2PaymentArranged: '2. Payment Arranged',
                step2Desc: 'We contact you for the $50 fee',
                step3ReviewBegins: '3. Review Begins',
                step3Desc: 'After payment, we review your application',
                propertyHeader: 'Property & Applicant Details',
                propertyInfo: 'Property Information',
                propertyAddressLabel: 'Property Address Applying For',
                propertyAddressPlaceholder: 'Street address, city, state, zip',
                errAddress: 'Please enter the property address',
                moveInLabel: 'Requested Move-in Date',
                errRequired: 'Required',
                leaseTermLabel: 'Desired Lease Term',
                selectTerm: 'Select term...',
                months6: '6 Months',
                months12: '12 Months',
                months18: '18 Months',
                months24: '24 Months',
                monthToMonth: 'Month-to-month',
                primaryApplicantInfo: 'Primary Applicant Information',
                firstNameLabel: 'First Name',
                lastNameLabel: 'Last Name',
                emailLabel: 'Email Address',
                emailPlaceholder: 'email@example.com',
                emailHint: 'Make sure the provided email is correct and accessible. Confirmation and updates sent here.',
                errEmail: 'Invalid email',
                phoneLabel: 'Phone Number',
                phonePlaceholder: '(555) 000-0000',
                phoneHint: 'Our team will contact you here.',
                errPhone: 'Invalid phone',
                dobLabel: 'Date of Birth',
                ssnLabel: 'Social Security Number (Last 4 Digits)',
                ssnHint: 'Only last 4 digits required',
                ssnPlaceholder: '1234',
                coApplicantCheckbox: 'I have a co-applicant or guarantor',
                coApplicantInfo: 'Additional Person Information',
                coRoleLabel: 'Role (Select one)',
                roleCoApplicant: 'Co-applicant (will live in the unit)',
                roleGuarantor: 'Guarantor (financial backup only)',
                coFirstNameLabel: 'First Name',
                coLastNameLabel: 'Last Name',
                coEmailLabel: 'Email',
                coPhoneLabel: 'Phone',
                coDobLabel: 'Date of Birth',
                coSsnLabel: 'SSN (Last 4)',
                employmentIncome: 'Employment & Income',
                coEmployerLabel: 'Employer',
                coJobTitleLabel: 'Job Title',
                coMonthlyIncomeLabel: 'Gross Monthly Income ($)',
                coMonthlyIncomePlaceholder: 'e.g., 4000',
                coEmploymentDurationLabel: 'Length of Employment',
                coEmploymentDurationPlaceholder: 'e.g., 2 years',
                coConsentLabel: 'I authorise verification of the information provided for this additional person, including credit and background check.',
                contactPrefsHeader: 'Contact Preferences (For Follow-up After Payment)',
                prefContactMethod: 'Preferred Contact Method',
                contactMethodText: 'Text Message',
                contactMethodEmail: 'Email',
                contactMethodHint: 'You can select both methods',
                availabilityLabel: 'Availability for Follow-up (After Payment)',
                weekdays: 'Weekdays',
                timeMorning: 'Morning (8am-11am)',
                timeMidday: 'Midday (11am-2pm)',
                timeAfternoon: 'Afternoon (2pm-5pm)',
                eveningsWeekends: 'Evenings & Weekends',
                timeEarlyEvening: 'Early Evening (5pm-8pm)',
                timeLateEvening: 'Late Evening (8pm-10pm)',
                timeWeekend: 'Weekend',
                flexible: 'Flexible',
                timeAnytime: 'Anytime — I\'m flexible',
                additionalNotesLabel: 'Additional Notes (Optional)',
                additionalNotesPlaceholder: 'e.g., Best after 7pm, avoid Wednesdays',
                preferencesNote: 'These preferences are for non-urgent follow-up after your payment is complete.',
                nextStep: 'Next Step',
                prevStep: 'Previous',
                editSection: 'Edit Section',
                residencyHeader: 'Residency & Occupancy',
                currentResidence: 'Current Residence',
                currentAddressLabel: 'Current Address',
                currentAddressPlaceholder: 'Street, Unit #, City, State, Zip',
                residencyStartLabel: 'How long at this address?',
                residencyStartPlaceholder: 'e.g., 2 years 3 months',
                rentAmountLabel: 'Current Rent/Mortgage Amount',
                rentAmountPlaceholder: '$',
                reasonLeavingLabel: 'Reason for leaving',
                landlordNameLabel: 'Current Landlord/Property Manager Name',
                landlordPhoneLabel: 'Landlord/Property Manager Phone',
                occupantsPets: 'Occupants & Pets',
                totalOccupantsLabel: 'Number of total occupants (including children)',
                occupantNamesLabel: 'Names and ages of all other occupants',
                occupantNamesPlaceholder: 'List names, ages, and relationship (e.g., Jane Doe, 7, daughter)',
                hasPetsLabel: 'Do you have any pets?',
                yes: 'Yes',
                no: 'No',
                petDetailsLabel: 'Pet details (type, breed, weight)',
                petDetailsPlaceholder: 'Describe your pets...',
                vehicleInfo: 'Vehicle Information',
                hasVehicleLabel: 'Do you have a vehicle?',
                vehicleMakeLabel: 'Make',
                vehicleModelLabel: 'Model',
                vehicleYearLabel: 'Year',
                vehicleYearPlaceholder: 'e.g., 2020',
                vehiclePlateLabel: 'License Plate (Optional)',
                employmentHeader: 'Employment & Income',
                currentEmployment: 'Current Employment',
                employmentStatusLabel: 'Employment Status',
                selectStatus: 'Select status...',
                fullTime: 'Full-time',
                partTime: 'Part-time',
                selfEmployed: 'Self-employed',
                student: 'Student',
                retired: 'Retired',
                unemployed: 'Unemployed',
                employerLabel: 'Employer',
                jobTitleLabel: 'Job Title',
                employmentDurationLabel: 'How long at this job?',
                employmentDurationPlaceholder: 'e.g., 3 years',
                supervisorNameLabel: 'Supervisor Name',
                supervisorPhoneLabel: 'Supervisor Phone',
                incomeVerification: 'Income Information',
                monthlyIncomeLabel: 'Gross Monthly Income',
                monthlyIncomePlaceholder: '$',
                incomeHint: 'Before taxes and deductions',
                otherIncomeLabel: 'Additional Monthly Income (Optional)',
                otherIncomePlaceholder: '$',
                otherIncomeHint: 'Child support, disability, etc.',
                incomeRatioLabel: 'Income-to-Rent Ratio:',
                financialHeader: 'References & Emergency Contact',
                personalReferences: 'Personal References',
                referencesHint: 'Please provide two references who are not related to you',
                ref1NameLabel: 'Reference 1 Name',
                ref1PhoneLabel: 'Reference 1 Phone',
                ref2NameLabel: 'Reference 2 Name (Optional)',
                ref2PhoneLabel: 'Reference 2 Phone (Optional)',
                emergencyInfo: 'Emergency Contact',
                emergencyNameLabel: 'Emergency Contact Name',
                emergencyPhoneLabel: 'Emergency Contact Phone',
                emergencyRelationshipLabel: 'Relationship to you',
                emergencyRelationshipPlaceholder: 'e.g., Spouse, Parent, Friend',
                additionalInfo: 'Additional Information',
                evictedLabel: 'Have you ever been evicted?',
                smokerLabel: 'Do you smoke?',
                paymentHeader: 'Payment Preferences',
                paymentIntro: 'Tell us which payment services you use. When we contact you about the $50 application fee, we\'ll discuss options you\'re familiar with.',
                paymentImportant: 'Payment must be completed before your application can be reviewed. Our team will contact you promptly after submission to arrange this.',
                primaryPref: 'Primary Preference',
                mainPaymentMethod: 'Your Main Payment Method',
                mainPaymentDesc: 'Which payment service do you use most often?',
                selectPrimary: '— Select your primary method —',
                other: 'Other',
                otherPaymentPlaceholder: 'Enter payment method',
                backupPref: 'Backup Options (Optional)',
                otherMethods: 'Other Methods You Use',
                otherMethodsDesc: 'If your primary isn\'t available, what else works for you?',
                secondaryMethod: 'Secondary Method',
                selectBackup: '— Select a backup (optional) —',
                thirdMethod: 'Third Method (Optional)',
                selectAnother: '— Select another (optional) —',
                duplicateWarning: 'Please select different payment methods for each choice.',
                reviewHeader: 'Review & Submit',
                feeTitle: 'Application Fee: $50.00',
                feeDesc: 'This fee is required before review can begin. Our team will contact you immediately after submission to arrange payment.',
                paymentReminderTitle: 'Payment Required Before Review',
                paymentReminderDesc: 'Your application is not complete until the $50 fee has been paid. Our team will contact you shortly after submission to arrange this.',
                verificationTitle: 'Verify Your Contact Information',
                verificationDesc: 'Please confirm your email and phone number are correct. This is how our team will reach you about the $50 fee.',
                reapplicationPolicyTextShort: 'If denied, apply again within 30 days with no new fee. Screening results valid for 60 days.',
                legalDeclaration: 'Legal Declaration',
                legalCertify: 'I certify that the information provided in this application is true and correct to the best of my knowledge.',
                legalAuthorize: 'I authorize verification of the information provided, including employment, income, and references.',
                termsAgreeLabel: 'I agree to the terms and conditions',
                submitBtn: 'Submit Application',
                submitDisclaimer: 'By clicking submit, your application will be securely transmitted to Choice Properties.',
                privacyPolicy: 'Privacy Policy',
                termsOfService: 'Terms of Service',
                contactSupport: 'Contact Support',
                progressSaved: 'Progress Saved',
                offlineMessage: 'You are currently offline. Progress will be saved locally.'
            },
            es: {
                langText: 'English',
                logoText: 'Choice Properties',
                tagline: 'Soluciones Profesionales de Administración de Propiedades',
                confidentialStamp: 'CONFIDENCIAL & SEGURO',
                trustIndicator: 'Su información está encriptada y protegida',
                timeEstimate: 'Tiempo estimado: 15-20 minutos',
                step1Label: 'Propiedad y Solicitante',
                step2Label: 'Residencia y Ocupación',
                step3Label: 'Empleo e Ingresos',
                step4Label: 'Finanzas y Referencias',
                step5Label: 'Preferencias de Pago',
                step6Label: 'Revisar y Enviar',
                stepPrefix: 'Paso',
                stepOf: 'de',
                processing: 'Procesando',
                validating: 'Validando',
                submitting: 'Enviando',
                complete: 'Completo',
                submittingTitle: 'Enviando su Solicitud',
                submissionMessage: 'Por favor no cierre esta ventana. Puede tomar unos momentos...',
                successTitle: 'Solicitud Recibida',
                successText: 'Gracias por elegir Choice Properties',
                appId: 'Su ID de Solicitud',
                clickToCopy: 'Copiar ID',
                immediateNextSteps: 'Próximos Pasos Inmediatos',
                paymentRequiredTitle: 'Pago Requerido Antes de la Revisión',
                paymentRequiredDesc: 'Nuestro equipo se comunicará con usted en breve al número proporcionado para coordinar el pago de $50.',
                completePaymentTitle: 'Completar el Pago',
                completePaymentDesc: 'Su solicitud no está completa hasta que se haya pagado la tarifa de $50. Discutiremos opciones de pago que conozca.',
                reviewBeginsTitle: 'Comienza la Revisión',
                reviewBeginsDesc: 'Una vez que se confirme el pago, su solicitud entra en el proceso de revisión formal. Puede seguir el estado en línea con su ID.',
                importantNote: 'Importante:',
                paymentUrgentText: 'Su solicitud no está completa hasta que se haya pagado la tarifa de $50. Por favor mantenga su teléfono cerca.',
                yourPreferences: 'Sus Preferencias (Para Seguimiento Después del Pago)',
                contactMethod: 'Método de Contacto:',
                bestTimes: 'Mejores Horarios:',
                paymentPref: 'Preferencias de Pago:',
                preferenceNote: 'Usaremos estas para seguimiento no urgente después de que se complete su pago.',
                questions: '¿Preguntas? Llame o envíe un mensaje de texto al',
                helpText: 'estamos aquí para ayudar.',
                trackStatus: 'Seguir Mi Solicitud',
                newApplication: 'Nueva Solicitud',
                reapplicationPolicyTitle: 'Protección de Reaplicación',
                reapplicationPolicyText: 'Si su solicitud es denegada, puede solicitar cualquier otra propiedad disponible dentro de los 30 días sin pagar otra tarifa de solicitud. Sus resultados de evaluación siguen siendo válidos por 60 días.',
                step1YouSubmit: '1. Usted Envía',
                step1Desc: 'Complete su solicitud completamente',
                step2PaymentArranged: '2. Pago Acordado',
                step2Desc: 'Lo contactamos para la tarifa de $50',
                step3ReviewBegins: '3. Comienza la Revisión',
                step3Desc: 'Después del pago, revisamos su solicitud',
                propertyHeader: 'Detalles de la Propiedad y el Solicitante',
                propertyInfo: 'Información de la Propiedad',
                propertyAddressLabel: 'Dirección de la Propiedad que Solicita',
                propertyAddressPlaceholder: 'Calle, ciudad, estado, código postal',
                errAddress: 'Por favor ingrese la dirección de la propiedad',
                moveInLabel: 'Fecha de Mudanza Solicitada',
                errRequired: 'Obligatorio',
                leaseTermLabel: 'Plazo de Arrendamiento Deseado',
                selectTerm: 'Seleccionar plazo...',
                months6: '6 Meses',
                months12: '12 Meses',
                months18: '18 Meses',
                months24: '24 Meses',
                monthToMonth: 'Mes a mes',
                primaryApplicantInfo: 'Información del Solicitante Principal',
                firstNameLabel: 'Nombre',
                lastNameLabel: 'Apellido',
                emailLabel: 'Correo Electrónico',
                emailPlaceholder: 'email@ejemplo.com',
                emailHint: 'Asegúrese de que el correo proporcionado sea correcto y accesible. La confirmación y actualizaciones se enviarán aquí.',
                errEmail: 'Correo inválido',
                phoneLabel: 'Número de Teléfono',
                phonePlaceholder: '(555) 000-0000',
                phoneHint: 'Nuestro equipo lo contactará aquí.',
                errPhone: 'Teléfono inválido',
                dobLabel: 'Fecha de Nacimiento',
                ssnLabel: 'Número de Seguro Social (Últimos 4 dígitos)',
                ssnHint: 'Solo últimos 4 dígitos requeridos',
                ssnPlaceholder: '1234',
                coApplicantCheckbox: 'Tengo un co-solicitante o fiador',
                coApplicantInfo: 'Información de Persona Adicional',
                coRoleLabel: 'Rol (Seleccione uno)',
                roleCoApplicant: 'Co-solicitante (vivirá en la unidad)',
                roleGuarantor: 'Fiador (solo respaldo financiero)',
                coFirstNameLabel: 'Nombre',
                coLastNameLabel: 'Apellido',
                coEmailLabel: 'Correo Electrónico',
                coPhoneLabel: 'Teléfono',
                coDobLabel: 'Fecha de Nacimiento',
                coSsnLabel: 'SSN (Últimos 4)',
                employmentIncome: 'Empleo e Ingresos',
                coEmployerLabel: 'Empleador',
                coJobTitleLabel: 'Puesto',
                coMonthlyIncomeLabel: 'Ingreso Mensual Bruto ($)',
                coMonthlyIncomePlaceholder: 'ej., 4000',
                coEmploymentDurationLabel: 'Tiempo en el empleo',
                coEmploymentDurationPlaceholder: 'ej., 2 años',
                coConsentLabel: 'Autorizo la verificación de la información proporcionada para esta persona adicional, incluyendo verificación de crédito y antecedentes.',
                contactPrefsHeader: 'Preferencias de Contacto (Para Seguimiento Después del Pago)',
                prefContactMethod: 'Método de Contacto Preferido',
                contactMethodText: 'Mensaje de Texto',
                contactMethodEmail: 'Correo Electrónico',
                contactMethodHint: 'Puede seleccionar ambos métodos',
                availabilityLabel: 'Disponibilidad para Seguimiento (Después del Pago)',
                weekdays: 'Días de semana',
                timeMorning: 'Mañana (8am-11am)',
                timeMidday: 'Mediodía (11am-2pm)',
                timeAfternoon: 'Tarde (2pm-5pm)',
                eveningsWeekends: 'Tardes y Fines de Semana',
                timeEarlyEvening: 'Temprano en la tarde (5pm-8pm)',
                timeLateEvening: 'Tarde noche (8pm-10pm)',
                timeWeekend: 'Fin de semana',
                flexible: 'Flexible',
                timeAnytime: 'En cualquier momento — soy flexible',
                additionalNotesLabel: 'Notas Adicionales (Opcional)',
                additionalNotesPlaceholder: 'ej., Mejor después de las 7pm, evitar miércoles',
                preferencesNote: 'Usaremos estas para seguimiento no urgente después de que se complete su pago.',
                nextStep: 'Siguiente Paso',
                prevStep: 'Anterior',
                editSection: 'Editar Sección',
                residencyHeader: 'Residencia y Ocupación',
                currentResidence: 'Residencia Actual',
                currentAddressLabel: 'Dirección Actual',
                currentAddressPlaceholder: 'Calle, Número, Ciudad, Estado, Código Postal',
                residencyStartLabel: '¿Cuánto tiempo en esta dirección?',
                residencyStartPlaceholder: 'ej., 2 años 3 meses',
                rentAmountLabel: 'Monto Actual de Alquiler/Hipoteca',
                rentAmountPlaceholder: '$',
                reasonLeavingLabel: 'Razón para mudarse',
                landlordNameLabel: 'Nombre del Propietario/Administrador Actual',
                landlordPhoneLabel: 'Teléfono del Propietario/Administrador',
                occupantsPets: 'Ocupantes y Mascotas',
                totalOccupantsLabel: 'Número total de ocupantes (incluyendo niños)',
                occupantNamesLabel: 'Nombres y edades de todos los demás ocupantes',
                occupantNamesPlaceholder: 'Lista de nombres, edades y relación (ej., Juan Pérez, 7, hijo)',
                hasPetsLabel: '¿Tiene mascotas?',
                yes: 'Sí',
                no: 'No',
                petDetailsLabel: 'Detalles de la mascota (tipo, raza, peso)',
                petDetailsPlaceholder: 'Describa sus mascotas...',
                vehicleInfo: 'Información del Vehículo',
                hasVehicleLabel: '¿Tiene vehículo?',
                vehicleMakeLabel: 'Marca',
                vehicleModelLabel: 'Modelo',
                vehicleYearLabel: 'Año',
                vehicleYearPlaceholder: 'ej., 2020',
                vehiclePlateLabel: 'Placa (Opcional)',
                employmentHeader: 'Empleo e Ingresos',
                currentEmployment: 'Empleo Actual',
                employmentStatusLabel: 'Estado de Empleo',
                selectStatus: 'Seleccionar estado...',
                fullTime: 'Tiempo completo',
                partTime: 'Medio tiempo',
                selfEmployed: 'Trabajador independiente',
                student: 'Estudiante',
                retired: 'Jubilado',
                unemployed: 'Desempleado',
                employerLabel: 'Empleador',
                jobTitleLabel: 'Puesto',
                employmentDurationLabel: '¿Cuánto tiempo en este trabajo?',
                employmentDurationPlaceholder: 'ej., 3 años',
                supervisorNameLabel: 'Nombre del supervisor',
                supervisorPhoneLabel: 'Teléfono del supervisor',
                incomeVerification: 'Información de Ingresos',
                monthlyIncomeLabel: 'Ingreso Mensual Bruto',
                monthlyIncomePlaceholder: '$',
                incomeHint: 'Antes de impuestos y deducciones',
                otherIncomeLabel: 'Otros Ingresos Mensuales (Opcional)',
                otherIncomePlaceholder: '$',
                otherIncomeHint: 'Pensión alimenticia, discapacidad, etc.',
                incomeRatioLabel: 'Relación Ingreso-Alquiler:',
                financialHeader: 'Finanzas y Referencias',
                personalReferences: 'Referencias Personales',
                referencesHint: 'Por favor proporcione dos referencias que no sean parientes',
                ref1NameLabel: 'Nombre de Referencia 1',
                ref1PhoneLabel: 'Teléfono de Referencia 1',
                ref2NameLabel: 'Nombre de Referencia 2 (Opcional)',
                ref2PhoneLabel: 'Teléfono de Referencia 2 (Opcional)',
                emergencyInfo: 'Contacto de Emergencia',
                emergencyNameLabel: 'Nombre de Contacto de Emergencia',
                emergencyPhoneLabel: 'Teléfono de Contacto de Emergencia',
                emergencyRelationshipLabel: 'Relación con usted',
                emergencyRelationshipPlaceholder: 'ej., Cónyuge, Padre, Amigo',
                additionalInfo: 'Información Adicional',
                evictedLabel: '¿Ha sido desalojado alguna vez?',
                smokerLabel: '¿Fuma?',
                paymentHeader: 'Preferencias de Pago',
                paymentIntro: 'Díganos qué servicios de pago utiliza. Cuando lo contactemos acerca de la tarifa de solicitud de $50, discutiremos opciones con las que esté familiarizado.',
                paymentImportant: 'El pago debe completarse antes de que su solicitud pueda ser revisada. Nuestro equipo lo contactará rápidamente después del envío para organizar esto.',
                primaryPref: 'Preferencia Principal',
                mainPaymentMethod: 'Su Método de Pago Principal',
                mainPaymentDesc: '¿Qué servicio de pago usa con más frecuencia?',
                selectPrimary: '— Seleccione su método principal —',
                other: 'Otro',
                otherPaymentPlaceholder: 'Ingrese método de pago',
                backupPref: 'Opciones de Respaldo (Opcional)',
                otherMethods: 'Otros Métodos Que Usa',
                otherMethodsDesc: 'Si su principal no está disponible, ¿qué más le funciona?',
                secondaryMethod: 'Método Secundario',
                selectBackup: '— Seleccione un respaldo (opcional) —',
                thirdMethod: 'Tercer Método (Opcional)',
                selectAnother: '— Seleccione otro (opcional) —',
                duplicateWarning: 'Por favor seleccione diferentes métodos de pago para cada opción.',
                reviewHeader: 'Revisar y Enviar',
                feeTitle: 'Tarifa de Solicitud: $50.00',
                feeDesc: 'Esta tarifa es requerida antes de que la revisión pueda comenzar. Nuestro equipo lo contactará inmediatamente después del envío para organizar el pago.',
                paymentReminderTitle: 'Pago Requerido Antes de la Revisión',
                paymentReminderDesc: 'Su solicitud no está completa hasta que se haya pagado la tarifa de $50. Nuestro equipo lo contactará poco después del envío para organizar esto.',
                verificationTitle: 'Verifique Su Información de Contacto',
                verificationDesc: 'Por favor confirme que su correo electrónico y número de teléfono sean correctos. Así es como nuestro equipo lo contactará acerca de la tarifa de $50.',
                reapplicationPolicyTextShort: 'Si es denegado, puede aplicar nuevamente dentro de 30 días sin nueva tarifa. Resultados de evaluación válidos por 60 días.',
                legalDeclaration: 'Declaración Legal',
                legalCertify: 'Certifico que la información proporcionada en esta solicitud es verdadera y correcta a mi leal saber y entender.',
                legalAuthorize: 'Autorizo la verificación de la información proporcionada, incluyendo empleo, ingresos y referencias.',
                termsAgreeLabel: 'Acepto los términos y condiciones',
                submitBtn: 'Enviar Solicitud',
                submitDisclaimer: 'Al hacer clic en enviar, su solicitud será transmitida de forma segura a Choice Properties.',
                privacyPolicy: 'Política de Privacidad',
                termsOfService: 'Términos de Servicio',
                contactSupport: 'Contactar Soporte',
                progressSaved: 'Progreso Guardado',
                offlineMessage: 'Actualmente está sin conexión. El progreso se guardará localmente.'
            }
        };

        this.translations = translations;
        this.state.language = 'en';
        const btn = document.getElementById('langToggle');
        const text = document.getElementById('langText');
        
        if (btn && text) {
            btn.addEventListener('click', () => {
                this.state.language = this.state.language === 'en' ? 'es' : 'en';
                const t = translations[this.state.language];
                text.textContent = t.langText;
                
                // Update all elements with data-i18n attribute
                document.querySelectorAll('[data-i18n]').forEach(el => {
                    const key = el.getAttribute('data-i18n');
                    if (t[key]) {
                        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                            if (el.placeholder !== undefined) el.placeholder = t[key];
                        } else if (el.tagName === 'OPTION') {
                            el.textContent = t[key];
                        } else {
                            el.textContent = t[key];
                        }
                    }
                });

                // Update buttons with icons (keep icon)
                document.querySelectorAll('.btn-next').forEach(b => {
                    const icon = b.querySelector('i');
                    const textSpan = document.createElement('span');
                    textSpan.setAttribute('data-i18n', 'nextStep');
                    textSpan.textContent = t.nextStep;
                    b.innerHTML = '';
                    b.appendChild(textSpan);
                    if (icon) b.appendChild(icon);
                });
                document.querySelectorAll('.btn-prev').forEach(b => {
                    const icon = b.querySelector('i');
                    const textSpan = document.createElement('span');
                    textSpan.setAttribute('data-i18n', 'prevStep');
                    textSpan.textContent = t.prevStep;
                    b.innerHTML = '';
                    if (icon) b.appendChild(icon);
                    b.appendChild(textSpan);
                });

                // Update progress bar text (which is not a data-i18n element)
                this.updateProgressBar();

                // If currently on step 6, regenerate the summary to translate "Edit Section"
                if (this.getCurrentSection() === 6) {
                    this.generateApplicationSummary();
                }

                this.saveProgress();
            });
        }
    }

    updateBilingualLabels(t) {
        // This method is kept for backward compatibility but no longer needed
        // as data-i18n handles all translations now.
    }

    updateLabel(id, text) {
        const label = document.querySelector(`label[for="${id}"]`);
        if (label) {
            const isRequired = label.classList.contains('required');
            label.textContent = text;
            if (isRequired) label.classList.add('required');
        }
    }

    // ---------- Generate summary ----------
    generateApplicationSummary() {
        const summaryContainer = document.getElementById('applicationSummary');
        if (!summaryContainer) return;

        const form = document.getElementById('rentalApplication');
        const formData = new FormData(form);
        const data = {};
        formData.forEach((value, key) => {
            if (value && key !== 'Application ID') {
                data[key] = value;
            }
        });

        const t = this.getTranslations(); // Get current translations

        const groups = [
            { id: 1, name: 'Property & Applicant', fields: [
                'Property Address', 'Requested Move-in Date', 'Desired Lease Term',
                'First Name', 'Last Name', 'Email', 'Phone', 'DOB', 'SSN'
            ]},
            { id: 1, name: 'Co-Applicant', fields: [
                'Has Co-Applicant', 'Additional Person Role',
                'Co-Applicant First Name', 'Co-Applicant Last Name',
                'Co-Applicant Email', 'Co-Applicant Phone',
                'Co-Applicant DOB', 'Co-Applicant SSN',
                'Co-Applicant Employer', 'Co-Applicant Job Title',
                'Co-Applicant Monthly Income', 'Co-Applicant Employment Duration',
                'Co-Applicant Consent'
            ]},
            { id: 2, name: 'Residency', fields: [
                'Current Address', 'Residency Duration', 'Current Rent Amount',
                'Reason for leaving', 'Current Landlord Name', 'Landlord Phone'
            ]},
            { id: 2, name: 'Occupancy & Vehicles', fields: [
                'Total Occupants', 'Additional Occupants', 'Has Pets', 'Pet Details',
                'Has Vehicle', 'Vehicle Make', 'Vehicle Model', 'Vehicle Year', 'Vehicle License Plate'
            ]},
            { id: 3, name: 'Employment & Income', fields: [
                'Employment Status', 'Employer', 'Job Title', 'Employment Duration',
                'Supervisor Name', 'Supervisor Phone', 'Monthly Income', 'Other Income'
            ]},
            { id: 4, name: 'Financial & References', fields: [
                'Emergency Contact Name', 'Emergency Contact Phone', 'Emergency Contact Relationship',
                'Reference 1 Name', 'Reference 1 Phone', 'Reference 2 Name', 'Reference 2 Phone'
            ]},
            { id: 5, name: 'Payment Preferences', fields: [
                'Primary Payment Method', 'Primary Payment Method Other',
                'Alternative Payment Method', 'Alternative Payment Method Other',
                'Third Choice Payment Method', 'Third Choice Payment Method Other'
            ]}
        ];

        const displayLabels = {
            'SSN': 'SSN (Last 4 Digits)',
            'Co-Applicant SSN': 'Co-Applicant SSN (Last 4)',
            'Has Co-Applicant': 'Has Co-Applicant/Guarantor',
            'Additional Person Role': 'Role'
        };

        let summaryHtml = '';
        groups.forEach(group => {
            let groupFieldsHtml = '';
            group.fields.forEach(field => {
                const value = data[field];
                const displayLabel = displayLabels[field] || field;
                if (value && value !== '') {
                    groupFieldsHtml += `
                        <div class="summary-item">
                            <div class="summary-label">${displayLabel}</div>
                            <div class="summary-value">${value}</div>
                        </div>`;
                }
            });

            if (groupFieldsHtml) {
                summaryHtml += `
                    <div class="summary-group" onclick="window.app.goToSection(${group.id})" style="cursor: pointer; transition: background 0.2s;">
                        <div class="summary-header">
                            <span>${group.name}</span>
                            <span style="font-size: 12px; color: var(--secondary); display: flex; align-items: center; gap: 4px;">
                                <i class="fas fa-edit"></i> ${t.editSection}
                            </span>
                        </div>
                        <div class="summary-content">
                            ${groupFieldsHtml}
                        </div>
                    </div>`;
            }
        });

        summaryContainer.innerHTML = summaryHtml;
    }

    goToSection(sectionNumber) {
        this.hideSection(this.getCurrentSection());
        this.showSection(sectionNumber);
        this.updateProgressBar();
    }

    // ---------- Handle form submission (connected to backend) ----------
async handleFormSubmit(e) {
    e.preventDefault();

    // Validate all sections
    for (let i = 1; i <= 5; i++) {
    if (!this.validateStep(i)) {
        this.showSection(i);
        this.updateProgressBar();
        return;
    }
}

    const submitBtn = document.getElementById('mainSubmitBtn');
    if (submitBtn) {
        submitBtn.classList.add('loading');
        submitBtn.disabled = true;
    }

    this.setState({ isSubmitting: true });
    this.showSubmissionProgress();

    try {
        this.updateSubmissionProgress(1, 'Sending application...');

        // Collect form data
        const form = document.getElementById('rentalApplication');
        const formData = new FormData(form);

        // Add generated appId
        const timestamp = Date.now();
        const randomSuffix = Math.random().toString(36).substring(2, 8).toUpperCase();
        const applicationId = `CP-${timestamp}-${randomSuffix}`;
        formData.append('appId', applicationId);  // send to backend

        this.updateSubmissionProgress(2, 'Submitting to server...');
        const response = await fetch(this.BACKEND_URL, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            this.updateSubmissionProgress(3, 'Application received!');
            await this.delay(500);
            this.handleSubmissionSuccess(result.appId);
        } else {
            throw new Error(result.error || 'Submission failed');
        }

    } catch (error) {
        console.error('Submission error:', error);
        const submitBtn = document.getElementById('mainSubmitBtn');
        if (submitBtn) {
            submitBtn.classList.remove('loading');
            submitBtn.disabled = false;
        }
        this.setState({ isSubmitting: false });
        this.hideSubmissionProgress();
        alert(`Submission failed: ${error.message}`);
    }
}

    // ---------- Submission progress UI ----------
    showSubmissionProgress() {
        const progress = document.getElementById('submissionProgress');
        const form = document.getElementById('rentalApplication');
        if (progress) progress.style.display = 'block';
        if (form) form.style.display = 'none';
    }

    hideSubmissionProgress() {
        const progress = document.getElementById('submissionProgress');
        const form = document.getElementById('rentalApplication');
        if (progress) progress.style.display = 'none';
        if (form) form.style.display = 'block';
    }

    updateSubmissionProgress(step, message) {
        const indicators = document.querySelectorAll('.step-indicator');
        indicators.forEach((el, idx) => {
            el.classList.toggle('active', idx + 1 === step);
        });
        const msgEl = document.getElementById('submissionMessage');
        if (msgEl) msgEl.textContent = message;
    }

    handleSubmissionSuccess(appId) {
        this.hideSubmissionProgress();
        this.showSuccessState(appId);
        this.clearSavedProgress();
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
// ---------- Success page ----------
    showSuccessState(appId) {
        const successState = document.getElementById('successState');
        if (!successState) return;

        const t = this.getTranslations();
        
        const getSelectedCheckboxValues = (name) => {
            const checkboxes = document.querySelectorAll(`input[name="${name}"]:checked`);
            return Array.from(checkboxes).map(cb => cb.value);
        };
        
        const contactMethods = getSelectedCheckboxValues('Preferred Contact Method');
        const contactTimes = getSelectedCheckboxValues('Preferred Time');
        const primaryPayment = document.getElementById('primaryPayment')?.value || 'Not selected';
        const secondaryPayment = document.getElementById('secondaryPayment')?.value;
        const thirdPayment = document.getElementById('thirdPayment')?.value;
        
        let paymentPrefs = primaryPayment;
        if (secondaryPayment && secondaryPayment !== '') paymentPrefs += `, ${secondaryPayment}`;
        if (thirdPayment && thirdPayment !== '') paymentPrefs += `, ${thirdPayment}`;
        
        successState.style.display = 'block';
        successState.innerHTML = `
            <div class="success-card">
                <div class="success-header">
                    <i class="fas fa-check-circle"></i>
                    <h2>${t.successTitle}</h2>
                    <p class="success-subtitle">${t.successText}</p>
                </div>

                <div class="id-section">
                    <div class="id-label">${t.appId}</div>
                    <div class="id-number" id="successAppId">${appId}</div>
                    <button class="copy-btn" onclick="copyAppId()">
                        <i class="fas fa-copy"></i> ${t.clickToCopy}
                    </button>
                </div>

                <div class="divider"></div>

                <div class="next-steps-box">
                    <h3><i class="fas fa-clock"></i> ${t.immediateNextSteps}</h3>
                    
                    <div class="step-row">
                        <div class="step-number">1</div>
                        <div class="step-content">
                            <strong>${t.paymentRequiredTitle}</strong>
                            <p>${t.paymentRequiredDesc}</p>
                        </div>
                    </div>

                    <div class="step-row">
                        <div class="step-number">2</div>
                        <div class="step-content">
                            <strong>${t.completePaymentTitle}</strong>
                            <p>${t.completePaymentDesc}</p>
                        </div>
                    </div>

                    <div class="step-row">
                        <div class="step-number">3</div>
                        <div class="step-content">
                            <strong>${t.reviewBeginsTitle}</strong>
                            <p>${t.reviewBeginsDesc}</p>
                        </div>
                    </div>
                </div>

                <div class="urgent-notice">
                    <i class="fas fa-exclamation-circle"></i>
                    <p><strong>${t.importantNote}</strong> ${t.paymentUrgentText}</p>
                </div>

                <div class="preference-summary">
                    <h4><i class="fas fa-clipboard-list"></i> ${t.yourPreferences}</h4>
                    <div class="preference-grid">
                        <div class="pref-item">
                            <span class="pref-label">${t.contactMethod}</span>
                            <span class="pref-value">${contactMethods.length > 0 ? contactMethods.join(', ') : 'Not specified'}</span>
                        </div>
                        <div class="pref-item">
                            <span class="pref-label">${t.bestTimes}</span>
                            <span class="pref-value">${contactTimes.length > 0 ? contactTimes.join(', ') : 'Not specified'}</span>
                        </div>
                        <div class="pref-item">
                            <span class="pref-label">${t.paymentPref}</span>
                            <span class="pref-value">${paymentPrefs}</span>
                        </div>
                    </div>
                    <p class="pref-note">${t.preferenceNote}</p>
                </div>

                <div class="policy-box">
                    <i class="fas fa-gem"></i>
                    <div>
                        <strong>${t.reapplicationPolicyTitle}</strong>
                        <p>${t.reapplicationPolicyText}</p>
                    </div>
                </div>

                <div class="action-buttons">
                    <!-- UPDATED LINK TO STATIC DASHBOARD -->
                    <a href="https://yourdomain.com/applicant-dashboard.html?id=${appId}" class="btn-track">
                        <i class="fas fa-chart-line"></i> ${t.trackStatus}
                    </a>
                    <button onclick="location.reload()" class="btn-new">
                        <i class="fas fa-plus"></i> ${t.newApplication}
                    </button>
                </div>

                <div class="help-line">
                    ${t.questions} <strong>707-706-3137</strong> — ${t.helpText}
                </div>
            </div>
        `;
        
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    getTranslations() {
        return this.translations[this.state.language] || this.translations['en'];
    }

    clearSavedProgress() {
        localStorage.removeItem(this.config.LOCAL_STORAGE_KEY);
    }
}

// ---------- Global copy function ----------
window.copyAppId = function() {
    const appId = document.getElementById('successAppId').innerText;
    navigator.clipboard.writeText(appId);
    alert('Application ID copied to clipboard');
};

// ============================================================
// TEST DATA FILL FUNCTIONALITY - KEPT FOR DEVELOPMENT
// ============================================================
(function() {
    const initTestButton = () => {
        const testBtn = document.getElementById('testFillBtn');
        if (!testBtn) return;
        
        console.log('✅ Test button initialized');
        
        testBtn.addEventListener('click', function(e) {
            e.preventDefault();
            fillTestData();
        });
    };
    
    function fillTestData() {
        console.log('🧪 Filling test data...');
        
        const today = new Date();
        const futureDate = new Date();
        futureDate.setDate(today.getDate() + 30);
        const futureDateStr = futureDate.toISOString().split('T')[0];
        
        const pastDateStr = '1990-01-15';
        
        // Step 1
        safeSetValue('propertyAddress', '123 Main Street, Troy, MI 48083');
        safeSetValue('requestedMoveIn', futureDateStr);
        safeSetValue('firstName', 'John');
        safeSetValue('lastName', 'Testerson');
        safeSetValue('email', 'test@example.com');
        safeSetValue('phone', '(555) 123-4567');
        safeSetValue('dob', pastDateStr);
        safeSetValue('ssn', '1234');
        safeSetValue('contactTimeSpecific', 'Best after 6pm, avoid Wednesdays');
        
        safeSetSelect('desiredLeaseTerm', '12 months');
        
        safeSetCheckbox('contactMethodText', true);
        safeSetCheckbox('contactMethodEmail', true);
        
        safeSetCheckbox('timeMorning', true);
        safeSetCheckbox('timeMidday', false);
        safeSetCheckbox('timeAfternoon', false);
        safeSetCheckbox('timeEarlyEvening', true);
        safeSetCheckbox('timeLateEvening', false);
        safeSetCheckbox('timeWeekend', true);
        safeSetCheckbox('timeAnytime', false);
        
        // Co-applicant
        safeSetCheckbox('hasCoApplicant', true);
        safeSetCheckbox('roleCoApplicant', true);
        safeSetValue('coFirstName', 'Jane');
        safeSetValue('coLastName', 'Testerson');
        safeSetValue('coEmail', 'jane@example.com');
        safeSetValue('coPhone', '(555) 987-6543');
        safeSetValue('coDob', '1992-03-20');
        safeSetValue('coSsn', '5678');
        safeSetValue('coEmployer', 'ABC Corp');
        safeSetValue('coJobTitle', 'Analyst');
        safeSetValue('coMonthlyIncome', '4500');
        safeSetValue('coEmploymentDuration', '3 years');
        safeSetCheckbox('coConsent', true);
        
        // Step 2
        safeSetValue('currentAddress', '456 Oak Avenue, Troy, MI 48083');
        safeSetValue('residencyStart', '3 years 2 months');
        safeSetValue('rentAmount', '1500');
        safeSetValue('reasonLeaving', 'Relocating for work opportunity');
        safeSetValue('landlordName', 'Sarah Johnson');
        safeSetValue('landlordPhone', '(555) 987-6543');
        safeSetValue('totalOccupants', '2');
        safeSetValue('occupantNames', 'Emma (age 7, daughter)');
        
        document.getElementById('petsYes')?.click();
        safeSetValue('petDetails', 'One friendly golden retriever, 65 lbs');
        
        // Vehicle
        document.getElementById('vehicleYes')?.click();
        safeSetValue('vehicleMake', 'Toyota');
        safeSetValue('vehicleModel', 'Camry');
        safeSetValue('vehicleYear', '2020');
        safeSetValue('vehiclePlate', 'ABC123');
        
        // Step 3
        safeSetSelect('employmentStatus', 'Full-time');
        safeSetValue('employer', 'Tech Solutions Inc');
        safeSetValue('jobTitle', 'Project Manager');
        safeSetValue('employmentDuration', '2 years');
        safeSetValue('supervisorName', 'Michael Chen');
        safeSetValue('supervisorPhone', '(555) 456-7890');
        safeSetValue('monthlyIncome', '5500');
        safeSetValue('otherIncome', '500');
        
        // Step 4
        safeSetValue('ref1Name', 'Robert Miller');
        safeSetValue('ref1Phone', '(555) 222-3333');
        safeSetValue('ref2Name', 'Lisa Thompson');
        safeSetValue('ref2Phone', '(555) 444-5555');
        safeSetValue('emergencyName', 'Jane Testerson');
        safeSetValue('emergencyPhone', '(555) 666-7777');
        safeSetValue('emergencyRelationship', 'Spouse');
        
        document.getElementById('evictedNo')?.click();
        document.getElementById('smokeNo')?.click();
        
        // Step 5
        safeSetSelect('primaryPayment', 'Venmo');
        safeSetSelect('secondaryPayment', 'PayPal');
        safeSetSelect('thirdPayment', 'Cash App');
        
        document.querySelectorAll('.other-payment-input').forEach(el => {
            el.style.display = 'none';
        });
        
        // Step 6
        safeSetCheckbox('certifyCorrect', true);
        safeSetCheckbox('authorizeVerify', true);
        safeSetCheckbox('termsAgree', true);
        
        if (window.app && typeof window.app.saveProgress === 'function') {
            window.app.saveProgress();
        }
        
        showTestFillNotification();
        
        if (window.app && typeof window.app.showSection === 'function') {
            window.app.showSection(1);
            window.app.updateProgressBar();
        }
        
        console.log('✅ Test data filled successfully');
    }
    
    function safeSetValue(id, value) {
        const el = document.getElementById(id);
        if (el) {
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }
    }
    
    function safeSetSelect(id, value) {
        const el = document.getElementById(id);
        if (el) {
            el.value = value;
            el.dispatchEvent(new Event('change', { bubbles: true }));
            
            if (id === 'primaryPayment' && value === 'Other') {
                document.getElementById('primaryPaymentOtherContainer').style.display = 'block';
            }
            if (id === 'secondaryPayment' && value === 'Other') {
                document.getElementById('secondaryPaymentOtherContainer').style.display = 'block';
            }
            if (id === 'thirdPayment' && value === 'Other') {
                document.getElementById('thirdPaymentOtherContainer').style.display = 'block';
            }
        }
    }
    
    function safeSetCheckbox(id, checked) {
        const el = document.getElementById(id);
        if (el) {
            el.checked = checked;
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }
    
    function showTestFillNotification() {
        const existing = document.getElementById('testFillNotification');
        if (existing) existing.remove();
        
        const notification = document.createElement('div');
        notification.id = 'testFillNotification';
        notification.innerHTML = `
            <i class="fas fa-check-circle"></i>
            <span>Test data filled! You can now edit any field.</span>
        `;
        notification.style.cssText = `
            position: fixed;
            top: 24px;
            left: 50%;
            transform: translateX(-50%);
            background: #10b981;
            color: white;
            padding: 16px 24px;
            border-radius: 60px;
            font-weight: 500;
            box-shadow: 0 10px 25px -5px rgba(16, 185, 129, 0.5);
            z-index: 100000;
            display: flex;
            align-items: center;
            gap: 12px;
            animation: slideDown 0.3s ease;
            border: 2px solid white;
        `;
        
        if (!document.getElementById('testNotificationStyle')) {
            const style = document.createElement('style');
            style.id = 'testNotificationStyle';
            style.textContent = `
                @keyframes slideDown {
                    from {
                        opacity: 0;
                        transform: translate(-50%, -20px);
                    }
                    to {
                        opacity: 1;
                        transform: translate(-50%, 0);
                    }
                }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            if (notification) {
                notification.style.opacity = '0';
                notification.style.transition = 'opacity 0.3s';
                setTimeout(() => notification.remove(), 300);
            }
        }, 3000);
    }
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initTestButton);
    } else {
        initTestButton();
    }
})();

// ---------- Initialize app ----------
document.addEventListener('DOMContentLoaded', () => {
    window.app = new RentalApplication();
    const s1 = document.getElementById('section1');
    if (s1) s1.classList.add('active');
});