import os
import sys
import contextlib
import io
import tempfile

# PaddlePaddle 3.x + older OCR models: avoid OneDNN fused_conv2d crash
os.environ.setdefault('FLAGS_use_mkldnn', '0')
os.environ.setdefault('FLAGS_onednn', '0')
os.environ.setdefault('FLAGS_enable_pir_api', '0')
os.environ.setdefault('GLOG_minloglevel', '2')

# Windows consoles often use cp1252; force UTF-8 for Vietnamese text
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import cv2
import argparse
import numpy as np
from PIL import Image


FONT = './PaddleOCR/doc/fonts/latin.ttf'
IMG_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp'}
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
VIETOCR_CONFIG_DIR = os.path.join(PROJECT_ROOT, 'vietocr', 'config')
VIETOCR_WEIGHT_NAME = 'vgg_transformer.pth'
VIETOCR_WEIGHT_URLS = (
    # Hugging Face mirror (works when vocr.vn DNS fails)
    'https://huggingface.co/datasets/qadeptrai123/aic26-ocr/resolve/main/models/ppocrv5-vietocr/vietocr_vgg_transformer.pth',
    'https://vocr.vn/data/vietocr/vgg_transformer.pth',
)


def imread_unicode(img_path, flags=cv2.IMREAD_COLOR):
    """Read image paths that contain non-ASCII characters on Windows."""
    try:
        data = np.fromfile(str(img_path), dtype=np.uint8)
    except OSError:
        return None
    if data.size == 0:
        return None
    return cv2.imdecode(data, flags)


def imwrite_unicode(img_path, image):
    """Write image paths that contain non-ASCII characters on Windows."""
    extension = os.path.splitext(str(img_path))[1] or '.png'
    ok, encoded = cv2.imencode(extension, image)
    if not ok:
        return False
    encoded.tofile(str(img_path))
    return True


def collect_images(path):
    if os.path.isfile(path):
        return [path]
    if os.path.isdir(path):
        return sorted(
            os.path.join(path, name)
            for name in os.listdir(path)
            if os.path.splitext(name)[1].lower() in IMG_EXTS
        )
    raise FileNotFoundError(f'No image or directory found at: {path}')


def predict(recognitor, detector, img_path, save_path=None, padding=4, verbose=True, draw_output=True):
    img = imread_unicode(img_path)
    if img is None:
        if verbose:
            print(f'Skipping unreadable image: {img_path}')
        return [], []

    result = detector.ocr(img, cls=False, det=True, rec=False)
    if not result or not result[0]:
        if verbose:
            print(f'No text detected in: {img_path}')
        return [], []
    result = result[0]

    boxes = []
    for line in result:
        boxes.append([[int(line[0][0]), int(line[0][1])], [int(line[2][0]), int(line[2][1])]])
    boxes = boxes[::-1]

    for box in boxes:
        box[0][0] = box[0][0] - padding
        box[0][1] = box[0][1] - padding
        box[1][0] = box[1][0] + padding
        box[1][1] = box[1][1] + padding

    kept_boxes = []
    cropped_images = []
    for box in boxes:
        box[0][0] = max(0, box[0][0])
        box[0][1] = max(0, box[0][1])
        box[1][0] = min(img.shape[1], box[1][0])
        box[1][1] = min(img.shape[0], box[1][1])
        cropped_image = img[box[0][1]:box[1][1], box[0][0]:box[1][0]]
        if cropped_image.size == 0:
            continue
        try:
            cropped_image = Image.fromarray(cv2.cvtColor(cropped_image, cv2.COLOR_BGR2RGB))
        except Exception:
            continue

        cropped_images.append(cropped_image)
        kept_boxes.append(box)

    if hasattr(recognitor, 'predict_batch') and cropped_images:
        texts = recognitor.predict_batch(cropped_images)
    else:
        texts = [recognitor.predict(cropped_image) for cropped_image in cropped_images]

    if verbose:
        for text in texts:
            print(text)

    if not draw_output or not save_path:
        return kept_boxes, texts

    def get_rectangle_points(x1, y1, x2, y2):
        return [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]

    draw_boxes = [
        get_rectangle_points(box[0][0], box[0][1], box[1][0], box[1][1])
        for box in kept_boxes
    ]
    from PaddleOCR import draw_ocr

    out_img = draw_ocr(img, draw_boxes, texts, scores=None, font_path=FONT)

    os.makedirs(save_path, exist_ok=True)
    img_name = os.path.basename(img_path)
    imwrite_unicode(os.path.join(save_path, img_name), out_img)
    return kept_boxes, texts


def predict_fast(detector, img_path, save_path=None, verbose=True, draw_output=True):
    img = imread_unicode(img_path)
    if img is None:
        if verbose:
            print(f'Skipping unreadable image: {img_path}')
        return [], []

    result = detector.ocr(img, cls=False, det=True, rec=True)
    if not result or not result[0]:
        if verbose:
            print(f'No text detected in: {img_path}')
        return [], []

    boxes = []
    texts = []
    scores = []
    for line in result[0]:
        boxes.append(line[0])
        texts.append(line[1][0])
        scores.append(line[1][1])

    if verbose:
        for text in texts:
            print(text)

    if not draw_output or not save_path:
        return boxes, texts

    from PaddleOCR import draw_ocr

    out_img = draw_ocr(img, boxes, texts, scores=scores, font_path=FONT)
    os.makedirs(save_path, exist_ok=True)
    img_name = os.path.basename(img_path)
    imwrite_unicode(os.path.join(save_path, img_name), out_img)
    return boxes, texts


def create_fast_model(use_gpu_arg=None, quiet=False, det_limit_side_len=1600):
    use_gpu = str(use_gpu_arg).lower() in ('1', 'true', 'yes') if use_gpu_arg is not None else False
    with contextlib.ExitStack() as stack:
        if quiet:
            silent_output = io.StringIO()
            stack.enter_context(contextlib.redirect_stdout(silent_output))
            stack.enter_context(contextlib.redirect_stderr(silent_output))

        from PaddleOCR import PaddleOCR

        return PaddleOCR(
            use_angle_cls=False,
            lang='vi',
            use_gpu=use_gpu,
            enable_mkldnn=False,
            show_log=not quiet,
            det_limit_side_len=det_limit_side_len,
            det_db_thresh=0.2,
            det_db_box_thresh=0.35,
            det_db_unclip_ratio=1.8,
            cpu_threads=max(1, min(4, os.cpu_count() or 1)),
        )


def _is_valid_weight_file(path, min_bytes=1_000_000):
    try:
        return os.path.isfile(path) and os.path.getsize(path) >= min_bytes
    except OSError:
        return False


def resolve_vietocr_weights(quiet=False):
    """Prefer local weights; otherwise download from HF mirror / vocr.vn."""
    candidates = (
        os.path.join(PROJECT_ROOT, 'weights', VIETOCR_WEIGHT_NAME),
        os.path.join(tempfile.gettempdir(), VIETOCR_WEIGHT_NAME),
    )
    for path in candidates:
        if _is_valid_weight_file(path):
            return path

    target = candidates[0]
    os.makedirs(os.path.dirname(target), exist_ok=True)

    import requests
    from tqdm import tqdm

    last_error = None
    for url in VIETOCR_WEIGHT_URLS:
        partial = target + '.partial'
        try:
            if not quiet:
                print(f'Downloading VietOCR weights from {url}')
            with requests.get(url, stream=True, timeout=60) as response:
                response.raise_for_status()
                total = int(response.headers.get('content-length') or 0)
                with open(partial, 'wb') as output:
                    iterator = response.iter_content(chunk_size=8192)
                    if not quiet and total:
                        iterator = tqdm(iterator, total=total // 8192, unit='chunk')
                    for chunk in iterator:
                        if chunk:
                            output.write(chunk)
            if not _is_valid_weight_file(partial):
                raise RuntimeError(f'Downloaded file is too small: {partial}')
            os.replace(partial, target)
            return target
        except Exception as exc:
            last_error = exc
            if os.path.exists(partial):
                try:
                    os.remove(partial)
                except OSError:
                    pass
            if not quiet:
                print(f'Weight download failed ({url}): {exc}')

    raise ConnectionError(
        'Could not download VietOCR weights. Place the file at '
        f'{candidates[0]} or fix DNS/network access to Hugging Face / vocr.vn. '
        f'Last error: {last_error}'
    )


def load_local_vietocr_config(quiet=False):
    import yaml
    from vietocr.vietocr.tool.config import Cfg

    with open(os.path.join(VIETOCR_CONFIG_DIR, 'base.yml'), encoding='utf-8') as base_file:
        config = yaml.safe_load(base_file)
    with open(os.path.join(VIETOCR_CONFIG_DIR, 'vgg-transformer.yml'), encoding='utf-8') as model_file:
        config.update(yaml.safe_load(model_file))

    config['weights'] = resolve_vietocr_weights(quiet=quiet)
    return Cfg(config)


def create_models(use_gpu_arg=None, quiet=False, beamsearch=False, det_limit_side_len=1600):
    with contextlib.ExitStack() as stack:
        if quiet:
            silent_output = io.StringIO()
            stack.enter_context(contextlib.redirect_stdout(silent_output))
            stack.enter_context(contextlib.redirect_stderr(silent_output))

        import torch
        from PaddleOCR import PaddleOCR
        from vietocr.vietocr.tool.predictor import Predictor

        config = load_local_vietocr_config(quiet=quiet)
        config['cnn']['pretrained'] = False
        config['predictor']['beamsearch'] = beamsearch

        if use_gpu_arg is not None:
            use_gpu = str(use_gpu_arg).lower() in ('1', 'true', 'yes')
        else:
            use_gpu = torch.cuda.is_available()
        config['device'] = 'cuda' if use_gpu else 'cpu'
        torch.set_grad_enabled(False)
        if not use_gpu:
            torch.set_num_threads(max(1, min(4, os.cpu_count() or 1)))

        recognitor = Predictor(config)
        detector = PaddleOCR(
            use_angle_cls=False,
            lang='vi',
            use_gpu=use_gpu,
            enable_mkldnn=False,
            show_log=not quiet,
            rec=False,
            det_limit_side_len=det_limit_side_len,
            det_db_thresh=0.2,
            det_db_box_thresh=0.35,
            det_db_unclip_ratio=1.8,
            cpu_threads=max(1, min(4, os.cpu_count() or 1)),
        )

    return recognitor, detector


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--img', required=True, help='path to image file or directory')
    parser.add_argument('--output', default='./runs/predict', help='path to save output images')
    parser.add_argument('--use_gpu', required=False, help='use GPU? true/false')
    parser.add_argument('--quiet', action='store_true', help='hide model initialization logs')
    parser.add_argument('--beamsearch', action='store_true', help='use slower beam search recognition for higher accuracy')
    parser.add_argument('--no_draw', action='store_true', help='do not save image with detected OCR boxes')
    parser.add_argument('--det_side', type=int, default=1600, help='max side length for text detection; lower is faster')
    parser.add_argument('--engine', choices=('paddle', 'vietocr'), default='vietocr', help='vietocr keeps Vietnamese accents; paddle is faster but may lose accents')
    args = parser.parse_args()

    if args.engine == 'vietocr':
        recognitor, detector = create_models(args.use_gpu, args.quiet, args.beamsearch, args.det_side)
    else:
        recognitor = None
        detector = create_fast_model(args.use_gpu, args.quiet, args.det_side)

    image_paths = collect_images(args.img)
    if not image_paths:
        raise FileNotFoundError(f'No images found in: {args.img}')

    for img_path in image_paths:
        print(f'\n=== {img_path} ===')
        if args.engine == 'vietocr':
            predict(recognitor, detector, img_path, args.output, padding=6, draw_output=not args.no_draw)
        else:
            predict_fast(detector, img_path, args.output, draw_output=not args.no_draw)


if __name__ == '__main__':
    main()
