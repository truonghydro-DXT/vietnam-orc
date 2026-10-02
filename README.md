# Vietnamese OCR

This project is about Optical Character Recognition (OCR) in Vietnamese texts. It uses PaddleOCR and VietOCR frameworks to achieve this. PaddleOCR is a popular OCR framework that provides a wide range of OCR models and tools. VietOCR is a popular framework for Vietnamese OCR task, based on Transformer OCR architecture.

>Note that: this model is just a compiling model, which means that I have simply gathered scripts from models in order to create a cohesive and comprehensive result. The end-to-end project will be started in the near future.

# Outline

1. Text Detection
2. Text Recognition

# Text Dectection
Text detection is the process of locating text in an image or video and recognizing the presence of characters. The [DB algorithm](https://github.com/PaddlePaddle/PaddleOCR/blob/release/2.6/doc/doc_en/algorithm_det_db_en.md) is a popular algorithm used in the [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) framework to localize text in the input image. It works by detecting the text regions in the image and then grouping them into text lines. This algorithm is known for its high accuracy and speed.

To enhance the accuracy of Text Recognition, images cropped by the DB algorithm were padded. This is because the padding helps to ensure that the text is not cut off during the recognition process.

# Text Recognition

Text Recognition is the process of recognizing the text in an image or video. For Text Recognition part, you used [VietOCR](https://github.com/pbcquoc/vietocr), which is a popular framework for Vietnamese OCR task. It is based on Transformer OCR architecture. The Transformer OCR architecture is a combination of the CNN and Transformer models. The CNN model is used to extract features from the input image, while the Transformer model is used to recognize the text in the image. This architecture is known for its high accuracy and speed.

# Usage

Firstly, clone this repository by executing:

```
git clone https://github.com/bmd1905/vietnamese-ocr
```

After cloning the repository, download the required dependencies by running:

```
pip install -r requirement.txt
```

For command-line usage, execute the following script for inference:

```
.venv\Scripts\Activate.ps1
python ui.py

python -m pip install pymupdf

python predict.py
    --img path/to/image
    --output path/of/output_image
```
$env:PYTHONIOENCODING='utf-8'
python predict.py --img samples --output result

For Jupyter Notebook, you can explore and experiment with the code at [predict.ipynb](https://github.com/bmd1905/vietnamese-ocr/blob/master/predict.ipynb).

# Document APIs

Start the server with `python ui.py`, then call one of:

| API | Document |
|---|---|
| `POST /api/cccd-front` | Mặt trước CCCD |
| `POST /api/cccd-back` | Mặt sau CCCD |
| `POST /api/land-certificate` | Giấy CNSDĐ (Mẫu số 1) |
| `POST /api/land-certificate-2` | Giấy CNSDĐ (Mẫu số 2, **≥2 ảnh hoặc 1 PDF ≥2 trang**) |
| `POST /api/land-certificate-3` | Giấy CNSDĐ (Mẫu số 3) |

Each API loads image/PDF, runs OCR, and returns **exactly** the same JSON as UI Export JSON (`export_ocr.make_json` / `build_json`).

Multipart upload:

```powershell
curl.exe -X POST "http://127.0.0.1:8000/api/cccd-front" -F "file=@samples/can-cuoc.jpg" -F "engine=vietocr"
curl.exe -X POST "http://127.0.0.1:8000/api/cccd-front" -F "file=@samples/CC_truoc.jpg" -F "engine=vietocr"
curl.exe -X POST "http://127.0.0.1:8000/api/cccd-back" -F "file=@samples/CCCD_Sau.jpg" -F "engine=vietocr"
curl.exe -X POST "http://127.0.0.1:8000/api/cccd-back" -F "file=@samples/CCCD_Sau1.jpg" -F "engine=vietocr"
curl.exe -X POST "http://127.0.0.1:8000/api/land-certificate" -F "file=@samples/GCN-1.jpg" -F "engine=vietocr"

curl.exe -X POST "http://127.0.0.1:8000/api/land-certificate-3" -F "file=@samples/GCN-3.png" -F "engine=vietocr"

curl.exe -X POST "http://127.0.0.1:8000/api/land-certificate-2" -F "file=@samples/GCN_2-23.jpg" -F "file=@samples/GCN_2-14.jpg" -F "engine=vietocr"
# 1 PDF nhiều trang (≥ 2): -F "file=@document.pdf"

```
Response examples (same keys as Export JSON):

```json
// /api/cccd-front
{ "CitizenId": "", "OwnerName": "", "OwnerBirthYear": "", "OwnerGender": "", "Address": "" }

// /api/cccd-back
{ "Address": "", "OwnerIssueDate": "", "OwnerIssuePlace": "" }

// /api/land-certificate / land-certificate-2 / land-certificate-3
{
  "SheetNumber": "", "ParcelNumber": "", "Area": "", "ParcelAddress": "",
  "LandUse_GCN": "", "LandUse_Area": "", "OwnerCode": "", "LandType": "",
  "OwnerName": "", "CitizenId": ""
}
```

# References

- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)
- [VietOCR](https://github.com/pbcquoc/vietocr)
