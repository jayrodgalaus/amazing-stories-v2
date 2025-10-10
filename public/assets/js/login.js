
async function loginUser() {
    const email = document.getElementById('email').value.trim();
    const password = document.getElementById('password').value.trim();

    if (!email || !password) {
        alert('Please enter both email and password.');
        return;
    }

    try {
        const response = await fetch('/api/login', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ email, password })
        });

        const result = await response.json();

        if (response.ok) {
            console.log('Login response:', result);
            // Handle success (e.g., redirect, show message)
        } else {
            console.error('Login failed:', result);
            alert(result.error || 'Login failed');
        }
    } catch (error) {
        console.error('Error calling login API:', error);
        alert('Something went wrong.');
    }
}