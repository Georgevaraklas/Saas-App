import os
import requests


CLOUDMERSIVE_API_KEY = os.environ.get("CLOUDMERSIVE_API_KEY")


def scan_file_with_cloudmersive(file):

    url = "https://api.cloudmersive.com/virus/scan/file/advanced"

    headers = {
        "Apikey": CLOUDMERSIVE_API_KEY,
        "allowExecutables": "false",
        "allowInvalidFiles": "false",
        "allowScripts": "false",
        "allowPasswordProtectedFiles": "false",
        "allowMacros": "false",
        "allowXmlExternalEntities": "false",
        "restrictFileTypes": ".pdf"
    }

    files = {
        "inputFile": (
            file.filename,
            file.stream,
            file.content_type
        )
    }

    response = requests.post(
        url,
        headers=headers,
        files=files,
        timeout=30
    )

    response.raise_for_status()

    return response.json()




def scan_image_with_cloudmersive(file):

    url = "https://api.cloudmersive.com/virus/scan/file/advanced"

    headers = {
        "Apikey": CLOUDMERSIVE_API_KEY,
        "allowExecutables": "false",
        "allowInvalidFiles": "false",
        "allowScripts": "false",
        "allowPasswordProtectedFiles": "false",
        "allowMacros": "false",
        "allowXmlExternalEntities": "false",
        "restrictFileTypes": ".png,.jpg,.jpeg,.webp"
    }

    files = {
        "inputFile": (
            file.filename,
            file.stream,
            file.content_type
        )
    }

    response = requests.post(
        url,
        headers=headers,
        files=files,
        timeout=30
    )

    response.raise_for_status()

    return response.json()