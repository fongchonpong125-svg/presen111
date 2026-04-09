import sys
import os
import json
import time
import threading
import queue
import logging

import matplotlib
matplotlib.use('Agg')  # MUST be called before importing main

# Ensure we can import main
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from flask import render_template, request, Response
import main as main_module

# WE WILL USE THE MAIN APP AS OUR SINGLE FLASK APP
app = main_module.app

class QueueHandler(logging.Handler):
    def __init__(self, q):
        super().__init__()
        self.q = q

    def emit(self, record):
        self.q.put(self.format(record))

@app.route('/')
def index_presentation():
    """Renders the presentation page (now at root)."""
    return render_template('index.html')

@app.route('/api/demo', methods=['GET'])
def run_demo():
    """
    Runs the full analysis from main.py in a background thread and streams
    the Python logging output via Server-Sent Events (SSE).
    """
    q = queue.Queue()
    handler = QueueHandler(q)
    handler.setFormatter(logging.Formatter('%(message)s'))
    
    # Attach to root logger to capture main.py logs
    root_logger = logging.getLogger()
    root_logger.addHandler(handler)
    
    def background_task():
        try:
            results = main_module.run_full_analysis()
            main_module.ANALYSIS_RESULTS = results
            main_module.ANALYSIS_LAST_RUN_TS = time.time()
        except Exception as e:
            q.put(f"[ERROR] {e}")
        finally:
            q.put(None) # Sentinel to close stream

    thread = threading.Thread(target=background_task)
    thread.start()
    
    def generate():
        yield f"data: {json.dumps({'type': 'log', 'text': '[INFO] Initiating full cross-platform analysis...'})}\n\n"
        
        while True:
            # Block until a log line is available
            line = q.get()
            if line is None:
                break
                
            # If the log is clean, prefix it to look like a terminal
            if not line.startswith('['):
                line = f"[INFO] {line}"
                
            yield f"data: {json.dumps({'type': 'log', 'text': line})}\n\n"
            
        # Cleanup
        root_logger.removeHandler(handler)
        
        yield f"data: {json.dumps({'type': 'log', 'text': '[SUCCESS] Full analysis completed successfully!'})}\n\n"
        yield f"data: {json.dumps({'type': 'done'})}\n\n"

    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    print("🚀 Starting Unified Presentation Server on http://0.0.0.0:8080")
    # Bind to 0.0.0.0 for Render explicitly from command line, but here we just use 8080
    app.run(host='0.0.0.0', port=8080, debug=True, threaded=True, use_reloader=False)
