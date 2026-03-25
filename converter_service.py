"""Wraps coa_converter for background conversion."""

import os
import sys
import logging
import threading

logger = logging.getLogger(__name__)

# Resolve converter modules from local converter/ package
CONVERTER_DIR = os.path.join(os.path.dirname(__file__), 'converter')
if CONVERTER_DIR not in sys.path:
    sys.path.insert(0, CONVERTER_DIR)

from coa_converter import convert_coa  # noqa: E402
from supplier_checker import check_supplier  # noqa: E402


def check_needs_verification(pdf_path: str) -> dict:
    """Check if PDF is from a known supplier."""
    try:
        return check_supplier(pdf_path)
    except Exception as e:
        logger.error(f'Supplier check failed: {e}')
        return {'known': False, 'needs_ai_verification': True, 'message': str(e)}


def run_conversion(job_manager, job_id: str, pdf_path: str,
                   template_path: str, output_path: str,
                   on_complete=None):
    """Run conversion in a background thread."""

    def _convert():
        try:
            job_manager.update_job(job_id, status='converting')
            result_path = convert_coa(pdf_path, template_path, output_path)
            job_manager.update_job(job_id, output_path=result_path)

            # Check supplier status
            supplier_info = check_needs_verification(pdf_path)
            needs_verify = supplier_info.get('needs_ai_verification', True)
            job_manager.update_job(job_id, needs_verification=needs_verify)

            job = job_manager.get_job(job_id)
            force = job.get('force_verify', False)

            # Mark as converted (downloadable) immediately
            job_manager.update_job(job_id, status='converted')

            if needs_verify or force:
                job_manager.update_job(job_id, status='verifying')
                if on_complete:
                    on_complete(job_id, pdf_path, template_path, result_path)
            else:
                job_manager.update_job(job_id, status='done')

            # Keep input PDF for re-verification and multi-template conversion

        except Exception as e:
            logger.error(f'Conversion failed for job {job_id}: {e}')
            job_manager.update_job(job_id, status='error', error=str(e))

    t = threading.Thread(target=_convert, daemon=True)
    t.start()
    return t
