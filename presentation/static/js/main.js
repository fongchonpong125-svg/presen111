document.addEventListener('DOMContentLoaded', () => {
    // 1. Smooth Scrolling
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const targetId = this.getAttribute('href');
            if(targetId === '#') return;
            const targetEl = document.querySelector(targetId);
            if(targetEl) {
                targetEl.scrollIntoView({
                    behavior: 'smooth'
                });
            }
        });
    });

    // 2. Intersection Observer for fade-up animations
    const observerOptions = {
        threshold: 0.1,
        rootMargin: "0px 0px -50px 0px"
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
                observer.unobserve(entry.target);
            }
        });
    }, observerOptions);

    document.querySelectorAll('.fade-up').forEach(el => {
        observer.observe(el);
    });

    // 3. Terminal Demo Logic
    const runBtn = document.getElementById('run-btn');
    const termOutput = document.getElementById('term-output');
    const demoResult = document.getElementById('demo-result');

    function appendTerminal(text, type = 'info') {
        const line = document.createElement('div');
        line.className = 'term-line';
        if (text.includes('[ERROR]')) {
            line.classList.add('term-err');
        } else if (text.includes('[INFO]') || text.includes('[SUCCESS]')) {
            line.classList.add('term-info');
        }
        
        line.textContent = `> ${text}`;
        termOutput.appendChild(line);
        termOutput.scrollTop = termOutput.scrollHeight;
    }

    runBtn.addEventListener('click', () => {
        // Reset state
        termOutput.innerHTML = '';
        demoResult.classList.add('hidden');
        
        runBtn.disabled = true;
        runBtn.innerHTML = '<span class="icon">⌛</span> 正在执行...';
        
        appendTerminal(`Connecting to execution server...`, 'info');

        // Setup EventSource for SSE
        const source = new EventSource(`/api/demo`);
        
        source.onmessage = function(event) {
            const data = JSON.parse(event.data);
            
            if (data.type === 'log') {
                appendTerminal(data.text);
            } else if (data.type === 'ping') {
                // 心跳包，不显示在终端，仅用于维持连接
                console.log('Keep-alive heartbeat received');
            } else if (data.type === 'done') {
                source.close();
                runBtn.disabled = false;
                runBtn.innerHTML = '<span class="icon">▶</span> 执行全网分析脚本 (Run Scripts)';
                
                demoResult.classList.remove('hidden');
                setTimeout(() => {
                    demoResult.classList.add('visible');
                }, 100);
            }
        };

        source.onerror = function(err) {
            source.close();
            runBtn.disabled = false;
            runBtn.innerHTML = '<span class="icon">▶</span> 执行全网分析脚本 (Run Scripts)';
            appendTerminal('[ERROR] Server connection closed. The process may still be running in the background.', 'err');
            appendTerminal('请稍等15秒后，直接尝试点击下方出现的“打开仪表盘”按钮查看结果。', 'info');
            
            // 即便报错，也强行显示结果按钮，因为后端可能已经执行成功只是连接断了
            demoResult.classList.remove('hidden');
            demoResult.classList.add('visible');
        };
    });
});
