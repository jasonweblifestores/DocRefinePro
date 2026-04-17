# SAVE AS: docrefine/reporting.py
import json
from pathlib import Path
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader
from .config import SystemUtils

def generate_job_report(ws_path, action_name, file_results=None):
    try:
        ws = Path(ws_path)
        rpt_dir = ws / "04_Reports"
        rpt_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        file_name = f"Audit_Certificate_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        
        s = {}
        stats_path = ws / "stats.json"
        if stats_path.exists():
            with open(stats_path) as f: 
                s = json.load(f)
        
        total_orig = 0
        total_new = 0
        errors = []
        skipped = 0
        
        if file_results:
            for res in file_results:
                if res.get('skipped'): 
                    skipped += 1
                    continue
                total_orig += res.get('orig_size', 0)
                total_new += res.get('new_size', 0)
                if not res.get('ok', True):
                    errors.append({'file': res.get('file', '?'), 'msg': res.get('error', 'Unknown')})
        
        saved_bytes = total_orig - total_new
        saved_mb = round(saved_bytes / (1024 * 1024), 2)
        saved_pct = round((saved_bytes / total_orig * 100), 1) if total_orig > 0 else 0

        # Calculate breakdown times
        t_batch = str(timedelta(seconds=int(s.get('batch_time', 0))))
        files_processed = len(file_results) if file_results else s.get('total_scanned', 0)

        # Jinja setup
        template_dir = SystemUtils.get_resource_dir() / "docrefine" / "templates"
        
        # Load environment
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        template = env.get_template('report.html')
        
        html_out = template.render(
            version=SystemUtils.CURRENT_VERSION,
            action_name=action_name,
            timestamp=timestamp,
            files_processed=files_processed,
            skipped_count=skipped,
            saved_mb=saved_mb,
            saved_pct=saved_pct,
            error_count=len(errors),
            batch_duration=t_batch,
            errors=errors
        )
        
        with open(rpt_dir / file_name, "w", encoding="utf-8") as f:
            f.write(html_out)
            
        return str(rpt_dir / file_name)
    except Exception as e:
        print(f"Report Gen Error: {e}")
        return None
