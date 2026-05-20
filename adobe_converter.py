from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult


def convert_pdf_to_docx(pdf_bytes: bytes, client_id: str, client_secret: str) -> bytes:
    """
    Converts PDF bytes to DOCX bytes using Adobe PDF Services API.
    Returns the DOCX file as bytes.
    """
    try:
        # Set up credentials
        credentials = ServicePrincipalCredentials(
            client_id=client_id,
            client_secret=client_secret
        )

        # Create PDF Services instance
        pdf_services = PDFServices(credentials=credentials)

        # Upload the PDF bytes to Adobe
        input_asset = pdf_services.upload(
            input_stream=pdf_bytes,
            mime_type=PDFServicesMediaType.PDF
        )

        # Set export target format to DOCX
        export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)

        # Create and submit the conversion job
        export_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_params)
        location = pdf_services.submit(export_job)

        # Poll for result
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)

        # Read result as bytes and return
        result_asset = pdf_services_response.get_result().get_asset()
        stream_asset: StreamAsset = pdf_services.get_content(result_asset)
        return stream_asset.get_input_stream().read()

    except ServiceUsageException as e:
        raise RuntimeError(f"Adobe API quota exceeded — you may have hit the 500/month free tier limit. ({e})")
    except ServiceApiException as e:
        raise RuntimeError(f"Adobe API error — check your credentials or the PDF file. ({e})")
    except SdkException as e:
        raise RuntimeError(f"Adobe SDK error. ({e})")
