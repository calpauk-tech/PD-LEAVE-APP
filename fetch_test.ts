async function main() {
  const res = await fetch('https://openapi.planday.com/swagger/hr/swagger.json');
  const data = await res.json();
  console.log(Object.keys(data.paths));
}
main();
