from azure.storage.file import FileService

file_service = FileService(account_name="node3storage", account_key='eN4wZTOidD3ID/8z3XVP7SX7HgQv14gBZu5bDorvWi7JP7v4VJzkUjKodKGRPYrEfmYDJ8jxzN719vDMPpL+Ww==')
#file_service = FileService(account_name="node3storage", account_key="sv=2017-11-09&ss=f&srt=sco&sp=rwdlc&se=2018-12-02T01:07:38Z&st=2018-08-01T18:07:38Z&spr=https&sig=ZH%2FS8N0KmhKlxISxmoZBvRapW6XMCQU5wrEQiTFN6ik%3D")
file_service.get_file_to_path("brief-files/root/OMS", None, 'propuesta22.xlsx', 'propuestadescarga.xlsx')