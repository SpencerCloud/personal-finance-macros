function getDateFromIso8601String(iso8601String) {
  return new Date(iso8601String);
}

function convertDateToIso8601(date) {
  // Use Etc/GMT no matter what time zone you are in, so converting back and forth is correct
  return Utilities.formatDate(date, 'Etc/GMT', 'YYYY-MM-dd');
}

function getYesterday() {
  const MILLISECONDS_PER_DAY = 1000 * 60 * 60 * 24;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // Weird thing to help with avoiding UTC witchcraft
  let yesterday = new Date(today.getTime() - MILLISECONDS_PER_DAY);
  const yesterdayIso8601 = convertDateToIso8601(yesterday);
  yesterday = getDateFromIso8601String(yesterdayIso8601);

  return yesterday;
}


















