function isValidExtension(filepath) {
  const fileExtension = filepath.split('.').pop().toLowerCase();
  return ['xlsx', 'xls'].includes(fileExtension);
}

export { isValidExtension };
