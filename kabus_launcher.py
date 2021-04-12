import argparse
import logging
import subprocess
import sys
import time

import psutil
import pywinauto
import requests

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
_f_handler = logging.FileHandler('kabus_launcher.log', 'a+')
_f_handler.setLevel(logging.INFO)
logger.addHandler(_f_handler)
_s_handler = logging.StreamHandler(sys.stdout)
_s_handler.setLevel(logging.INFO)
logger.addHandler(_s_handler)


def run_dfsvc_if_needed():
    _process = [p for p in psutil.process_iter(attrs=["pid", "name"]) if p.info['name'] == "dfsvc.exe"]

    if not _process:
        logger.info('launch dfsvc.')
        subprocess.Popen("C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\dfsvc.exe", shell=True)
    else:
        logger.info('no launch dfsvc.')


def exit_kabus_exe_if_needed():
    _process = [p for p in psutil.process_iter(attrs=["pid", "name"]) if p.info['name'] == "KabuS.exe"]
    if _process:
        for _p in _process:
            logger.info(f'Terminated {_p.pid}')
            _p.terminate()
        gone, alive = psutil.wait_procs(_process, timeout=10)
        for _p in alive:
            logger.info('Killed {_p.pid}')
            _p.kill()


def waiting_for_kabus_api(retry=100, sleep=1):
    cnt = 0
    while cnt < retry:
        try:
            _r = requests.get("http://localhost:18081/kabusapi")
            return
        except:
            pass
        finally:
            cnt += 1
            time.sleep(sleep)


def launch(id, pwd, retry=100, sleep=1):
    logger.info(f'Launching kabuS.exe')
    cnt = 0
    while cnt < retry:
        try:
            exit_kabus_exe_if_needed()

            _p = subprocess.Popen([
                'C:\\Windows\\System32\\rundll32.exe',
                'C:\\Windows\\System32\\dfshim.dll,ShOpenVerbApplication',
                'http://download.r10.kabu.co.jp/kabustation/KabuStation.application'], shell=True)
            app = pywinauto.Application().connect(path='KabuS.exe', timeout=90)
            app['ログイン'].wait('ready', timeout=10)

            window = app['ログイン']
            window.Edit2.set_edit_text(id)
            window.Edit.set_edit_text(pwd)
            window.child_window(auto_id="btnOK").click()
            break
        except:
            logger.exception('Error At launch')
        finally:
            cnt += 1
            time.sleep(sleep)

    logger.info('waiting for kabus_api...')
    waiting_for_kabus_api()
    logger.info('Launched')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('id')
    parser.add_argument('pwd')
    args = parser.parse_args()

    run_dfsvc_if_needed()

    launch(args.id, args.pwd)
