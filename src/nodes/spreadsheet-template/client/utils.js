import { isValidExtension as validate } from '../utils';

function validateFilepathInput(value, typeElmId) {
  const type = $(typeElmId).val();
  if (type !== 'str') return true;
  return validate(value);
}

export { validateFilepathInput };
