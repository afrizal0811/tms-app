import requests
from utils.function import show_error_message
from utils.messages import ERROR_MESSAGES

def handle_requests_error(err):
    if isinstance(err, requests.exceptions.HTTPError):
        status_code = err.response.status_code
        if status_code == 401:
            show_error_message("Akses Ditolak (401)", ERROR_MESSAGES["API_TOKEN_MISSING"])
        elif status_code >= 500:
            show_error_message("Masalah Server API", ERROR_MESSAGES["SERVER_ERROR"].format(error_detail=status_code))
        else:
            show_error_message("Kesalahan HTTP", ERROR_MESSAGES["HTTP_ERROR_GENERIC"].format(status_code=status_code))

    elif isinstance(err, requests.exceptions.Timeout):
        show_error_message("Waktu Habis", ERROR_MESSAGES["TIMEOUT"])

    elif isinstance(err, requests.exceptions.TooManyRedirects):
        show_error_message("Redirect Berlebihan", ERROR_MESSAGES["TOO_MANY_REDIRECTS"])

    elif isinstance(err, requests.exceptions.ConnectionError):
        show_error_message("Koneksi Gagal", ERROR_MESSAGES["CONNECTION_ERROR"])

    elif isinstance(err, requests.exceptions.RequestException):
        show_error_message("Kesalahan API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=err))

    else:
        show_error_message("Kesalahan Tidak Dikenal", str(err))