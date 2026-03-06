document.addEventListener('DOMContentLoaded', () => {
    const recordBtn = document.getElementById('record-btn');
    const statusText = document.getElementById('status-text');
    const recordingIndicator = document.getElementById('recording-indicator');
    const resultForm = document.getElementById('result-form');
    const recorderContainer = document.getElementById('recorder-container');
    const errorMessage = document.getElementById('error-message');
    const errorText = document.getElementById('error-text');

    const amountInput = document.getElementById('amount');
    const descInput = document.getElementById('description');
    const yearInput = document.getElementById('year'); // New
    // Removed categoryInput
    const expenseTypeInput = document.getElementById('expense_type');

    const saveBtn = document.getElementById('save-btn');
    const cancelBtn = document.getElementById('cancel-btn');
    const successScreen = document.getElementById('success-screen');
    const newRecordBtn = document.getElementById('new-record-btn');

    let mediaRecorder;
    let audioChunks = [];
    let isRecording = false;

    const t = (key) => window.translations[key] || key;

    function resetUI() {
        resultForm.classList.add('hidden');
        successScreen.classList.add('hidden');
        recorderContainer.classList.remove('hidden');
        errorMessage.classList.add('hidden');
        statusText.textContent = t('status_ready');
        statusText.classList.remove('animate-pulse');
        recordBtn.classList.remove('bg-red-500');
        recordBtn.classList.add('bg-brand-500');
        recordBtn.innerHTML = '<i class="fa-solid fa-microphone"></i>';
        recordingIndicator.classList.add('hidden');
        amountInput.value = '';
        descInput.value = '';
        yearInput.value = ''; // Reset year
        expenseTypeInput.value = 'Essential';
    }

    recordBtn.addEventListener('click', async () => {
        if (!isRecording) {
            try {
                const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
                mediaRecorder = new MediaRecorder(stream);
                audioChunks = [];

                mediaRecorder.ondataavailable = event => {
                    audioChunks.push(event.data);
                };

                mediaRecorder.onstop = uploadAudio;

                mediaRecorder.start();
                isRecording = true;

                recordBtn.classList.remove('bg-brand-500');
                recordBtn.classList.add('bg-red-500');
                recordBtn.innerHTML = '<i class="fa-solid fa-stop"></i>';
                statusText.textContent = t('status_recording');
                recordingIndicator.classList.remove('hidden');

            } catch (err) {
                console.error("Error accessing microphone", err);

                // Check for insecure origin issue on mobile
                if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1' && window.location.protocol !== 'https:') {
                    showError("Error accessing microphone: Not Secure. Enable 'Insecure origins treated as secure' in chrome://flags for this IP.");
                    alert("To use Microphone on Mobile without HTTPS:\n1. Go to chrome://flags\n2. Search 'Insecure origins treated as secure'\n3. Enable it and add: http://" + window.location.host + "\n4. Relaunch Chrome.");
                } else {
                    showError("Error accessing microphone: " + err.message);
                }
            }
        } else {
            mediaRecorder.stop();
            isRecording = false;

            recordBtn.classList.remove('bg-red-500');
            recordBtn.classList.add('bg-brand-500');
            recordBtn.innerHTML = '<i class="fa-solid fa-microphone"></i>';
            statusText.textContent = t('status_processing');
            statusText.classList.add('animate-pulse');
            recordingIndicator.classList.add('hidden');
        }
    });

    async function uploadAudio() {
        const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
        const formData = new FormData();
        formData.append('audio_data', audioBlob, 'voice.webm');

        try {
            const response = await fetch('/upload_audio', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) throw new Error('Network response was not ok');

            const data = await response.json();

            if (data.error) {
                showError(data.error);
                statusText.textContent = t('status_ready');
                return;
            }

            recorderContainer.classList.add('hidden');
            statusText.textContent = "";
            resultForm.classList.remove('hidden');

            amountInput.value = data.amount || '';
            descInput.value = data.description || '';
            yearInput.value = data.year || ''; // Populate year
            // No category input to populate

            if (data.expense_type && (data.expense_type === 'Essential' || data.expense_type === 'Side')) {
                expenseTypeInput.value = data.expense_type;
            } else {
                expenseTypeInput.value = 'Essential';
            }

        } catch (error) {
            console.error('Error:', error);
            showError("Processing Error.");
            resetUI();
        }
    }

    saveBtn.addEventListener('click', async () => {
        const amount = amountInput.value;
        const description = descInput.value;
        const expense_type = expenseTypeInput.value;
        const year = yearInput.value; // Get year

        if (!amount || !description) {
            showError("Fields incomplete");
            return;
        }

        saveBtn.disabled = true;
        saveBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin mx-2"></i> ${t('saving')}`;

        try {
            const response = await fetch('/save_expense', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ amount, description, expense_type, year }), // Send year
            });

            const data = await response.json();

            if (data.success) {
                resultForm.classList.add('hidden');
                successScreen.classList.remove('hidden');
            } else {
                showError("Save failed");
            }
        } catch (error) {
            showError("Connection error");
        } finally {
            saveBtn.disabled = false;
            saveBtn.innerHTML = `<i class="fa-solid fa-check mx-2"></i> ${t('save')}`;
        }
    });

    cancelBtn.addEventListener('click', resetUI);
    newRecordBtn.addEventListener('click', resetUI);

    function showError(msg) {
        errorText.textContent = msg;
        errorMessage.classList.remove('hidden');
        setTimeout(() => {
            errorMessage.classList.add('hidden');
        }, 5000);
    }
});
