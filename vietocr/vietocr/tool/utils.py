import os
import gdown
import yaml
import numpy as np
import uuid
import requests
import tempfile
from tqdm import tqdm

# Fallback mirrors when the primary host (often vocr.vn) is unreachable.
WEIGHT_URL_FALLBACKS = {
    'vgg_transformer.pth': (
        'https://huggingface.co/datasets/qadeptrai123/aic26-ocr/resolve/main/models/ppocrv5-vietocr/vietocr_vgg_transformer.pth',
        'https://vocr.vn/data/vietocr/vgg_transformer.pth',
    ),
    'vietocr_vgg_transformer.pth': (
        'https://huggingface.co/datasets/qadeptrai123/aic26-ocr/resolve/main/models/ppocrv5-vietocr/vietocr_vgg_transformer.pth',
        'https://vocr.vn/data/vietocr/vgg_transformer.pth',
    ),
}


def download_weights(uri, cached=None, md5=None, quiet=False):
    if uri.startswith('http'):
        return download(url=uri, quiet=quiet)
    return uri


def _candidate_urls(url):
    filename = url.split('/')[-1].split('?')[0]
    fallbacks = WEIGHT_URL_FALLBACKS.get(filename, ())
    urls = [url]
    for candidate in fallbacks:
        if candidate not in urls:
            urls.append(candidate)
    return urls, filename


def download(url, quiet=False):
    tmp_dir = tempfile.gettempdir()
    urls, filename = _candidate_urls(url)
    # Always cache under the canonical VietOCR filename when possible.
    if filename.startswith('vietocr_') and filename.endswith('.pth'):
        cache_name = filename[len('vietocr_'):]
    else:
        cache_name = filename
    full_path = os.path.join(tmp_dir, cache_name)

    if os.path.exists(full_path) and os.path.getsize(full_path) > 1_000_000:
        print('Model weight {} exsits. Ignore download!'.format(full_path))
        return full_path

    last_error = None
    for candidate in urls:
        partial = full_path + '.partial'
        try:
            with requests.get(candidate, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(partial, 'wb') as f:
                    for chunk in tqdm(r.iter_content(chunk_size=8192), disable=quiet):
                        if chunk:
                            f.write(chunk)
            os.replace(partial, full_path)
            return full_path
        except Exception as exc:
            last_error = exc
            if os.path.exists(partial):
                try:
                    os.remove(partial)
                except OSError:
                    pass

    raise ConnectionError(
        'Failed to download model weights from any mirror. Last error: {}'.format(last_error)
    )


def download_config(id):
    url = 'https://vocr.vn/data/vietocr/config/{}'.format(id)
    r = requests.get(url)
    config = yaml.safe_load(r.text)
    return config


def compute_accuracy(ground_truth, predictions, mode='full_sequence'):
    """
    Computes accuracy
    :param ground_truth:
    :param predictions:
    :param display: Whether to print values to stdout
    :param mode: if 'per_char' is selected then
                 single_label_accuracy = correct_predicted_char_nums_of_single_sample / single_label_char_nums
                 avg_label_accuracy = sum(single_label_accuracy) / label_nums
                 if 'full_sequence' is selected then
                 single_label_accuracy = 1 if the prediction result is exactly the same as label else 0
                 avg_label_accuracy = sum(single_label_accuracy) / label_nums
    :return: avg_accuracy
    """
    if mode == 'per_char':

        accuracy = []

        for index, label in enumerate(ground_truth):
            prediction = predictions[index]
            total_count = len(label)
            correct_count = 0
            try:
                for i, tmp in enumerate(label):
                    if tmp == prediction[i]:
                        correct_count += 1
            except IndexError:
                continue
            finally:
                try:
                    accuracy.append(correct_count / total_count)
                except ZeroDivisionError:
                    if len(prediction) == 0:
                        accuracy.append(1)
                    else:
                        accuracy.append(0)
        avg_accuracy = np.mean(np.array(accuracy).astype(np.float32), axis=0)
    elif mode == 'full_sequence':
        try:
            correct_count = 0
            for index, label in enumerate(ground_truth):
                prediction = predictions[index]
                if prediction == label:
                    correct_count += 1
            avg_accuracy = correct_count / len(ground_truth)
        except ZeroDivisionError:
            if not predictions:
                avg_accuracy = 1
            else:
                avg_accuracy = 0
    else:
        raise NotImplementedError('Other accuracy compute mode has not been implemented')

    return avg_accuracy
