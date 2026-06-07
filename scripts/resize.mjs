// 생성된 이미지를 정확한 780×2000 규격으로 후처리한다.
// 사용법: node scripts/resize.mjs <input.png> <output.png>
// resvg가 아닌 단순 비율 보정이 필요하면 추후 sharp 도입 예정.
import { readFileSync, writeFileSync } from 'node:fs';
import { Resvg } from '@resvg/resvg-js';

const [, , input, output] = process.argv;
if (!input || !output) {
  console.error('usage: node scripts/resize.mjs <input.png> <output.png>');
  process.exit(1);
}

const W = 780;
const H = 2000;
const b64 = readFileSync(input).toString('base64');
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
  <image href="data:image/png;base64,${b64}" x="0" y="0" width="${W}" height="${H}" preserveAspectRatio="xMidYMid slice"/>
</svg>`;
const png = new Resvg(svg, { fitTo: { mode: 'width', value: W } }).render().asPng();
writeFileSync(output, png);
console.log(`wrote ${output} (${W}x${H})`);
