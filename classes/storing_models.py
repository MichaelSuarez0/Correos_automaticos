from pydantic import BaseModel, EmailStr, RootModel
from typing import Optional, Literal

class EmailData(BaseModel):
    msg_id: str
    from_name: str
    from_email: EmailStr
    sent: Optional[str]
    to: EmailStr
    subject: str
    body: str
    attachments: list[str]


class AttachmentLog(BaseModel):
    new_name: str
    original_name: str
    path: str
    sharepoint_uploaded: Optional[bool] = None


"""
ic| email_data: {'1002': {'attachments': ['T80-REV_Indicador_Corrupcion_Modulo 85 '
                                          '(Recuperado).xlsx',
                                          'T80-Incremento de la '
                                          'corrupción_270623_MSP_JP_OS.docx',
                                          't80 - Incremento de la corrupción.do'],
                          'body': '',
                          'from_email': 'msuarez@ceplan.gob.pe',
                          'from_name': 'Michael Salvador Suarez Patilongo',
                          'sent': '',
                          'subject': 'Sistematizar',
                          'to': 'Consulta Técnica <consultatecnica@ceplan.gob.pe>'},

ic| user_attachments_log: {'msuarez@ceplan.gob.pe': [{'new_name': 't80 - Incremento de la '
                                                                  'corrupción.xlsx',
                                                      'original_name': 'T80-REV_Indicador_Corrupcion_Modulo '
                                                                       '85 (Recuperado).xlsx',
                                                      'path': 'Tendencias/Tendencias Nacionales',
                                                      'sharepoint_status': 'uploaded'},
                                                     {'new_name': 't80 - Incremento de la '
                                                                  'corrupción.docx',
                                                      'original_name': 'T80-Incremento de la '
                                                                       'corrupción_270623_MSP_JP_OS.docx',
                                                      'path': 'Tendencias/Tendencias Nacionales',
                                                      'sharepoint_status': 'uploaded'},
"""