class OlcumOlustur {
    constructor() {
        this.currentStep = 1;
        this.totalSteps = 2;
        this.bacaSayisi = 0;
        this.formData = {};
        this.parameters = [];
        
        // Initialize the app when DOM is loaded
        document.addEventListener('DOMContentLoaded', () => this.init());
    }
    
    init() {
        // Get DOM elements
        this.nextStepBtn = document.getElementById('nextStep');
        this.prevStepBtn = document.getElementById('prevStep');
        this.finishStepBtn = document.getElementById('finishStep');
        this.progressBar = document.getElementById('progressBar');
        this.currentStepSpan = document.getElementById('currentStep');
        this.stepDescription = document.getElementById('stepDescription');
        this.olcumBilgileri = document.getElementById('olcumBilgileri');
        this.bacaContainer = document.getElementById('bacaContainer');
        
        // Load parameters from datalist
        this.loadParameters();
        
        // Set up event listeners
        this.setupEventListeners();
        
        // Initialize form with default values
        this.initializeForm();
        
        // Update UI for the first step
        this.updateStepUI();
    }

    // Event listeners
    setupEventListeners() {
        // Next button
        if (this.nextStepBtn) {
            this.nextStepBtn.addEventListener('click', () => this.nextStep());
        }
        
        // Previous button
        if (this.prevStepBtn) {
            this.prevStepBtn.addEventListener('click', () => this.prevStep());
        }
        
        // Finish button
        if (this.finishStepBtn) {
            this.finishStepBtn.addEventListener('click', () => this.finishStep());
        }
    }
    
    // Load parameters from datalist
    loadParameters() {
        const parametreOptions = document.querySelectorAll('#parametreListesi option');
        this.parameters = Array.from(parametreOptions).map(opt => opt.value);
    }
    
    // Next step
    nextStep() {
        if (this.currentStep === 1) {
            if (this.validateStep1()) {
                this.collectStep1Data();
                this.bacaSayisi = this.formData.bacaSayisi;
                this.prepareStep2();
                this.currentStep++;
                this.updateStepUI();
            }
        }
    }

    // Previous step
    prevStep() {
        if (this.currentStep > 1) {
            this.currentStep--;
            this.updateStepUI();
        }
    }

    // Finish step
    finishStep() {
        if (this.validateStep2()) {
            this.submitForm();
        }
    }

    // Collect data from step 1
    collectStep1Data() {
        this.formData = {
            firmaAdi: document.getElementById('firmaAdi').value,
            olcumKodu: document.getElementById('olcumKodu').value,
            bacaSayisi: parseInt(document.getElementById('bacaSayisi').value),
            olcumBas: document.getElementById('olcumBas').value,
            olcumBit: document.getElementById('olcumBit').value,
            personeller: []
        };
        
        // Get selected personnel
        document.querySelectorAll('input[name="olcumPersoneli"]:checked').forEach(checkbox => {
            this.formData.personeller.push(checkbox.value);
        });
    }

    // Validate step 1
    validateStep1() {
        let isValid = true;
        const requiredFields = ['firmaAdi', 'olcumKodu', 'bacaSayisi', 'olcumBas', 'olcumBit'];
        
        requiredFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (!field || !field.value.trim()) {
                field?.classList.add('is-invalid');
                isValid = false;
            } else {
                field?.classList.remove('is-invalid');
            }
        });

        // At least one personnel must be selected
        if (document.querySelectorAll('input[name="olcumPersoneli"]:checked').length === 0) {
            this.showError('Lütfen en az bir personel seçin.');
            return false;
        }

        return isValid;
    }

    // Validate step 2
    validateStep2() {
        let isValid = true;
        
        for (let i = 1; i <= this.bacaSayisi; i++) {
            const bacaAdi = document.getElementById(`bacaAdi_${i}`);
            if (!bacaAdi || !bacaAdi.value.trim()) {
                bacaAdi?.classList.add('is-invalid');
                isValid = false;
            } else {
                bacaAdi?.classList.remove('is-invalid');
            }
            
            // At least one parameter must be selected
            if (document.querySelectorAll(`input[name="parametreler_${i}"]:checked`).length === 0) {
                this.showError(`Lütfen ${bacaAdi?.value || i}. baca için en az bir parametre seçin.`);
                return false;
            }
        }
        
        return isValid;
    }

    // Prepare step 2 with baca forms
    prepareStep2() {
        if (!this.bacaContainer) return;
        
        this.bacaContainer.innerHTML = '';
        
        // Add a header for the baca section
        const headerDiv = document.createElement('div');
        headerDiv.className = 'text-center mb-4';
        headerDiv.innerHTML = `
            <h5 class="text-primary">
                <i class="fas fa-industry me-2"></i>
                Baca ve Parametre Matrisi
            </h5>
            <p class="text-muted">Aşağıda her baca için ad ve ölçülecek parametreleri belirleyin</p>
        `;
        this.bacaContainer.appendChild(headerDiv);
        
        // Create a form for each baca
        for (let i = 1; i <= this.bacaSayisi; i++) {
            const bacaCard = document.createElement('div');
            bacaCard.className = 'baca-card';
            bacaCard.innerHTML = `
                <h6 class="mb-3">Baca ${i}</h6>
                <div class="row">
                    <div class="col-md-6">
                        <div class="mb-3">
                            <label for="bacaAdi_${i}" class="form-label">Baca Adı</label>
                            <input type="text" class="form-control" id="bacaAdi_${i}" required>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label">Ölçülecek Parametreler</label>
                        <div class="border rounded p-2" style="max-height: 200px; overflow-y: auto;">
                            ${this.parameters.map((param, idx) => `
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" 
                                        name="parametreler_${i}" 
                                        value="${param}" 
                                        id="param_${i}_${idx}">
                                    <label class="form-check-label" for="param_${i}_${idx}">${param}</label>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                </div>
            `;
            this.bacaContainer.appendChild(bacaCard);
        }
        
        // Update measurement info
        if (this.olcumBilgileri) {
            this.olcumBilgileri.textContent = `${this.formData.firmaAdi} - ${this.formData.olcumKodu}`;
        }
    }

    // Update UI based on current step
    updateStepUI() {
        // Hide all steps
        document.querySelectorAll('.step-content').forEach(step => {
            step.classList.remove('active');
        });
        
        // Show current step
        const currentStepElement = document.getElementById(`step${this.currentStep}`);
        if (currentStepElement) {
            currentStepElement.classList.add('active');
        }
        
        // Update progress bar
        const progressPercent = (this.currentStep / this.totalSteps) * 100;
        if (this.progressBar) {
            this.progressBar.style.width = `${progressPercent}%`;
            this.progressBar.setAttribute('aria-valuenow', progressPercent);
        }
        
        // Update step counter
        if (this.currentStepSpan) {
            this.currentStepSpan.textContent = this.currentStep;
        }
        
        // Update button states
        if (this.prevStepBtn) {
            this.prevStepBtn.style.display = this.currentStep === 1 ? 'none' : 'block';
        }
        
        if (this.nextStepBtn) {
            this.nextStepBtn.style.display = this.currentStep === this.totalSteps ? 'none' : 'block';
        }
        
        if (this.finishStepBtn) {
            this.finishStepBtn.style.display = this.currentStep === this.totalSteps ? 'block' : 'none';
        }
        
        // Update step description
        if (this.stepDescription) {
            this.stepDescription.textContent = this.currentStep === 1 ? 'Temel bilgiler' : 'Baca bilgileri ve parametreler';
        }
    }

    // Submit form data
    async submitForm() {
        try {
            // Collect baca data
            const bacalar = [];
            
            for (let i = 1; i <= this.bacaSayisi; i++) {
                const bacaAdi = document.getElementById(`bacaAdi_${i}`)?.value || `Baca ${i}`;
                const parametreler = [];
                
                document.querySelectorAll(`input[name="parametreler_${i}"]:checked`).forEach(checkbox => {
                    parametreler.push(checkbox.value);
                });
                
                bacalar.push({
                    bacaAdi,
                    parametreler
                });
            }
            
            // Combine all data
            const formDataToSubmit = {
                ...this.formData,
                bacalar
            };
            
            console.log('Form data to submit:', formDataToSubmit);
            
            // Here you would typically send the data to the server
            // Example:
            /*
            const response = await fetch('/api/olcum_olustur', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(formDataToSubmit)
            });
            
            if (!response.ok) {
                throw new Error('Sunucu hatası');
            }
            
            const result = await response.json();
            */
            
            // For now, just show success message
            this.showSuccess('Ölçüm başarıyla oluşturuldu!');
            
            // Close modal and refresh after delay
            setTimeout(() => {
                // In a real app, you might want to redirect or update the UI instead
                window.location.reload();
            }, 1500);
            
        } catch (error) {
            console.error('Error submitting form:', error);
            this.showError('Bir hata oluştu. Lütfen tekrar deneyin.');
        }
    }
    
    // Initialize form with default values
    initializeForm() {
        // Set default dates
        const today = new Date();
        const olcumBas = document.getElementById('olcumBas');
        if (olcumBas) {
            olcumBas.value = today.toISOString().slice(0, 16);
        }
        
        // Set default end time to 1 hour later
        const oneHourLater = new Date(today.getTime() + 60 * 60 * 1000);
        const olcumBit = document.getElementById('olcumBit');
        if (olcumBit) {
            olcumBit.value = oneHourLater.toISOString().slice(0, 16);
        }
        
        // Update plan tab when baca sayisi changes
        const bacaSayisiInput = document.getElementById('bacaSayisi');
        if (bacaSayisiInput) {
            bacaSayisiInput.addEventListener('change', (e) => {
                if (window.parent && window.parent.updateBacaSayisi) {
                    window.parent.updateBacaSayisi(e.target.value);
                }
            });
        }
    }
    
    // Show success message
    showSuccess(message) {
        if (window.Swal) {
            Swal.fire({
                icon: 'success',
                title: 'Başarılı',
                text: message,
                timer: 2000,
                showConfirmButton: false
            });
        } else {
            alert(message);
        }
    }
    
    // Show error message
    showError(message) {
        if (window.Swal) {
            Swal.fire({
                icon: 'error',
                title: 'Hata',
                text: message
            });
        } else {
            alert(message);
        }
    }
}

// Initialize the application
new OlcumOlustur();
