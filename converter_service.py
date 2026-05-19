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


def run_conversion(job_manager, job_id: str, pdf_path: str,
                   template_path: str, output_path: str,
                   on_complete=None):
    """Run conversion in a background thread."""

    def _convert():
        try:
            job_manager.update_job(job_id, status='converting')
            result_path = convert_coa(pdf_path, template_path, output_path)

            # Sanity check: result file must exist and be non-zero on disk.
            # On Windows we've seen reports of "blank download"; this log
            # confirms whether the conversion step actually wrote real bytes.
            if os.path.exists(result_path):
                size = os.path.getsize(result_path)
                logger.info(f'[转换] job {job_id} 输出文件: {result_path} ({size}B)')
                if size == 0:
                    logger.error(
                        f'[转换] job {job_id} 输出文件大小为 0！'
                        f'可能填充失败或文件被其他进程锁定。')
            else:
                logger.error(f'[转换] job {job_id} 输出文件不存在: {result_path}')

            job_manager.update_job(job_id, output_path=result_path)

            # Mark as converted (downloadable) immediately
            job_manager.update_job(job_id, status='converted')

            # Always proceed to AI verification
            job_manager.update_job(job_id, status='verifying')
            if on_complete:
                on_complete(job_id, pdf_path, template_path, result_path)

        except Exception as e:
            logger.error(f'Conversion failed for job {job_id}: {e}')
            job_manager.update_job(job_id, status='error', error=str(e))

    t = threading.Thread(target=_convert, daemon=True)
    t.start()
    return t
