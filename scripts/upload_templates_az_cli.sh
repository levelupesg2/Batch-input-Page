az storage blob upload-batch \
  --source "/path/to/Repo/Templates" \
  --destination "container-name" \
  --pattern "*.xlsx" \
  --account-name <storage-account-name>
