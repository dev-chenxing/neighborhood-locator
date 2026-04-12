const 街道映射 = require("./streets.json");

function getAddressNumber(street, address) {
  const regexp = new RegExp(`${street}(\\d+)号`);
  let number = address.match(regexp);
  if (number) return Number(number[1]);

  let matchResult = address.match(new RegExp(`${street}(.+)号`));
  if (matchResult) {
    number = matchResult[1];
    matchResult = number.match(/\d+/g);
  } else {
    matchResult = address.match(/\d+/g);
  }

  if (matchResult) {
    return Number(matchResult[0]);
  }

  return 0;
}

function 匹配所属社区(address) {
  for (const street in 街道映射) {
    const neighborhoods = 街道映射[street];
    if (!address.includes(street)) continue;

    for (const neighborhood of neighborhoods) {
      if (neighborhood["all"]) {
        return neighborhood["name"];
      }

      const addressNumber = getAddressNumber(street, address);
      if (neighborhood["oddity"] !== addressNumber % 2) {
        continue;
      }

      if (neighborhood["start"]) {
        if (neighborhood["start"] <= addressNumber && addressNumber <= neighborhood["end"]) {
          return neighborhood["name"];
        }
        continue;
      }

      return neighborhood["name"];
    }

    return undefined;
  }

  return undefined;
}

module.exports = {
  getAddressNumber,
  匹配所属社区,
};
