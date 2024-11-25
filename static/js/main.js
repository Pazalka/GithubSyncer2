document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const loadingSpinner = document.getElementById('loadingSpinner');
    const statusMessage = document.getElementById('statusMessage');
    const errorMessage = document.getElementById('errorMessage');

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        // Reset messages
        statusMessage.classList.add('d-none');
        errorMessage.classList.add('d-none');
        
        // Show loading state
        submitBtn.disabled = true;
        loadingSpinner.classList.remove('d-none');
        
        try {
            const formData = new FormData(form);
            
            // Collect all files
            let hasFiles = false;
            for (let i = 1; i <= 9; i++) {
                const fileInput = document.getElementById(`file${i}`);
                const files = fileInput.files;
                if (files && files.length > 0) {
                    hasFiles = true;
                    // Add all selected files to formData with unique names
                    Array.from(files).forEach((file, index) => {
                        formData.append(`file${i}_${index}`, file);
                    });
                }
            }

            if (!hasFiles) {
                throw new Error('אנא בחר לפחות קובץ אחד לעיבוד');
            }

            statusMessage.textContent = 'מעבד קבצים...';
            statusMessage.classList.remove('d-none');

            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'אירעה שגיאה בעיבוד הקבצים');
            }

            // Handle successful response
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'processed_output.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            statusMessage.textContent = 'העיבוד הושלם בהצלחה!';
            statusMessage.classList.remove('alert-info');
            statusMessage.classList.add('alert-success');

        } catch (error) {
            errorMessage.textContent = error.message;
            errorMessage.classList.remove('d-none');
        } finally {
            submitBtn.disabled = false;
            loadingSpinner.classList.add('d-none');
        }
    });
});
