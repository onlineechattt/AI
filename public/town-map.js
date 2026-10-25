// Lightweight, code-native miniature town; no extra image requests.
export const townMap = `<svg class="town-map" viewBox="0 0 1200 690" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
<defs>
 <linearGradient id="ground" x2="0" y2="1"><stop stop-color="#e6edda"/><stop offset="1" stop-color="#cddcc1"/></linearGradient>
 <linearGradient id="roof" x2=".8" y2="1"><stop stop-color="#9fbace"/><stop offset="1" stop-color="#6e92ad"/></linearGradient>
 <linearGradient id="wall" x2="0" y2="1"><stop stop-color="#fffdf2"/><stop offset="1" stop-color="#f0e2c9"/></linearGradient>
 <linearGradient id="glass" x2="1" y2="1"><stop stop-color="#c3dedc"/><stop offset=".48" stop-color="#e3eee7"/><stop offset="1" stop-color="#a8c5c6"/></linearGradient>
 <linearGradient id="canal"><stop stop-color="#b6d5d2"/><stop offset="1" stop-color="#8fbcb8"/></linearGradient>
 <linearGradient id="tree" x2="1" y2="1"><stop stop-color="#b8cba5"/><stop offset="1" stop-color="#779577"/></linearGradient>
 <filter id="soft-shadow" x="-40%" y="-40%" width="200%" height="200%"><feDropShadow dx="0" dy="12" stdDeviation="12" flood-color="#727f63" flood-opacity=".16"/></filter>
 <pattern id="pavers" width="30" height="20" patternUnits="userSpaceOnUse"><path d="M0 0H30V20H0Z M15 0V20" fill="#efe5d0" stroke="#e6dac1" stroke-width="1"/></pattern>
 <g id="mini-tree"><ellipse cx="7" cy="17" rx="29" ry="11" fill="#8a9e7524"/><path d="M0 10V-37" stroke="#aa9470" stroke-width="8" stroke-linecap="round"/><path d="M0-13l-12-16M0-20l12-15" stroke="#aa9470" stroke-width="5" stroke-linecap="round"/><path d="M0-92C-23-90-39-52-29-32C-20-9 23-12 31-34C40-56 18-91 0-92Z" fill="url(#tree)"/><ellipse cx="-9" cy="-60" rx="10" ry="19" fill="#d0dcba" opacity=".38"/></g>
 <g id="shrub"><ellipse rx="24" ry="10" fill="#859e7830"/><ellipse cy="-8" rx="21" ry="14" fill="#a9c296"/><ellipse cx="-9" cy="-12" rx="10" ry="11" fill="#bdd0a4"/></g>
 <g id="flower"><path d="M0 3V-9" stroke="#78966d" stroke-width="2"/><circle cy="-12" r="5" fill="#e8b4a2"/><circle cy="-12" r="2" fill="#fff0bf"/></g>
 <g id="lamp"><ellipse cy="8" rx="15" ry="5" fill="#82906f26"/><path d="M0 5V-67q0-9 9-9h9" fill="none" stroke="#829b93" stroke-width="5" stroke-linecap="round"/><path d="M9-72l-4 15h25l-5-15" fill="#f8e9bc" stroke="#829b93" stroke-width="3"/></g>
 <g id="box"><path d="M-19-9L0-19 21-8 1 3Z" fill="#dec69f"/><path d="M-19-9V11L1 24V3Z" fill="#c6aa7f"/><path d="M1 3V24L21 12V-8Z" fill="#b7986c"/><path d="M-8-15L12-3V6" fill="none" stroke="#f6e7cc" stroke-width="5"/></g>
 <g id="bench"><path d="M-28-5l57 0M-28-13h57M-28-21h57" stroke="#bd9e75" stroke-width="7" stroke-linecap="round"/><path d="M-20-18V14M20-18V14" stroke="#7f9582" stroke-width="5"/><path d="M-29 0h60" stroke="#d1b087" stroke-width="9" stroke-linecap="round"/></g>
</defs>
<!-- Floating model base and the little river. -->
<ellipse cx="610" cy="604" rx="509" ry="52" fill="#abb69d" opacity=".12"/>
<path d="M70 168Q70 136 103 136H1090Q1130 136 1130 176V546Q1130 572 1100 587L765 666Q741 673 714 668L94 561Q70 556 70 528Z" fill="#b6c6a6"/>
<path d="M70 150Q70 120 103 120H1090Q1130 120 1130 160V530Q1130 556 1100 571L765 650Q741 657 714 652L94 545Q70 540 70 512Z" fill="url(#ground)"/>
<path d="M929 129C914 228 1019 261 987 348S970 484 1115 526" fill="none" stroke="#d4e5d9" stroke-width="64"/>
<path d="M929 129C914 228 1019 261 987 348S970 484 1115 526" fill="none" stroke="url(#canal)" stroke-width="42"/>
<path d="M933 154q-8 16 0 30m38 87 6 19m-7 82-3 13m37 88 17 13" fill="none" stroke="#e5f0de" stroke-width="3" stroke-linecap="round" opacity=".8"/>
<!-- Walkways join the shop, warehouse and event booth. -->
<path d="M133 458C284 502 409 469 566 487S822 599 1058 546" fill="none" stroke="#c3c8ad" stroke-width="83"/>
<path d="M133 454C284 498 409 465 566 483S822 595 1058 542" fill="none" stroke="#f4ebd7" stroke-width="72"/>
<path d="M773 552Q862 486 857 337" fill="none" stroke="#f4ebd7" stroke-width="56"/>
<path d="M330 475L320 332M573 481V395" stroke="#f4ebd7" stroke-width="48"/>
<path d="M140 454C300 499 410 465 570 484S827 590 1056 542" fill="none" stroke="#dfd5bd" stroke-width="2" stroke-dasharray="3 13"/>
<!-- Park at left. -->
<path d="M128 267Q128 244 151 244H329Q351 244 351 266V373Q351 397 327 400L151 378Q128 375 128 350Z" fill="#a9be94"/>
<path d="M147 278Q147 265 161 265H315Q332 265 332 281V354Q332 371 314 371L162 354Q147 352 147 337Z" fill="#c8d7b0"/>
<ellipse cx="231" cy="315" rx="51" ry="31" fill="#eee6ce"/><ellipse cx="231" cy="308" rx="43" ry="25" fill="#9fc4be"/><ellipse cx="231" cy="304" rx="35" ry="18" fill="#bbd6cf"/>
<path d="M223 311q-8-20 5-24q16 3 11 21" fill="#fcf9e9"/><ellipse cx="232" cy="288" rx="10" ry="9" fill="#fffdf2"/><circle cx="235" cy="286" r="1.8" fill="#758d8f"/>
<use href="#bench" transform="translate(175 365) scale(.8)"/><use href="#bench" transform="translate(305 309) scale(.8)"/>
<use href="#mini-tree" transform="translate(144 279) scale(.85)"/><use href="#mini-tree" transform="translate(325 260) scale(.9)"/>
<use href="#shrub" transform="translate(155 393)"/><use href="#shrub" transform="translate(340 386)"/>
<!-- Warehouse, with a loading door and waiting parcels. -->
<g filter="url(#soft-shadow)"><path d="M777 207L867 172 953 207 865 245Z" fill="#b3bfad"/><path d="M777 207V331L865 371V245Z" fill="#e3d8bd"/><path d="M865 245V371L953 332V207Z" fill="#c6c4ad"/><path d="M767 207L866 160 963 206 865 253Z" fill="#9eafa0"/><path d="M767 207v10l98 47v-11ZM865 253v11l98-47v-11Z" fill="#879d90"/>
<path d="M799 250v76l48 22v-77Z" fill="#b6b6a2" stroke="#c4b797" stroke-width="5"/><path d="M803 265l40 18m-40-5 40 18m-40-5 40 18m-40-5 40 18m-40-5 40 18" stroke="#d5cfb8" stroke-width="3"/>
<path d="M884 271l49-21v37l-49 21Z" fill="url(#glass)" stroke="#f4eedb" stroke-width="5"/>
<path d="M796 229l52 24v17l-52-24Z" fill="#f9f6e9"/><text x="803" y="229" transform="rotate(25 803 229)" fill="#617a70" font-size="12" font-family="Georgia">WAREHOUSE</text></g>
<use href="#box" transform="translate(804 358) scale(.85)"/><use href="#box" transform="translate(829 374) scale(.8)"/>
<!-- The main lilbulle shop. -->
<ellipse cx="559" cy="425" rx="170" ry="31" fill="#87977c" opacity=".15"/>
<g filter="url(#soft-shadow)"><path d="M416 223L642 223 715 183V382L642 423H416Z" fill="url(#wall)"/>
<path d="M642 223L715 183V382L642 423Z" fill="#dccfb4"/>
<path d="M399 205L519 114 744 114 641 205Z" fill="url(#roof)"/><path d="M399 205v18h242v-18Z" fill="#597f9d"/><path d="M641 205L744 114v18l-103 91Z" fill="#54778d"/>
<path d="M426 190L530 114m-47 76 99-76m-38 76 97-76m-36 76 95-76" stroke="#c0d2dd" stroke-width="2" opacity=".48"/>
<path d="M514 117l9-13h224l6 10Z" fill="#b6cbd7"/>
<path d="M660 263l35-20v71l-35 19Z" fill="url(#glass)" stroke="#f8f3e5" stroke-width="5"/><path d="M678 252v72" stroke="#e8e7d5" stroke-width="4"/>
<rect x="433" y="286" width="116" height="102" rx="7" fill="#d4dfce" stroke="#d0bc98" stroke-width="7"/>
<rect x="563" y="287" width="58" height="131" rx="23" fill="url(#glass)" stroke="#d0bc98" stroke-width="7"/><path d="M591 293v120" stroke="#fcf6e7" stroke-width="5"/>
<circle cx="602" cy="364" r="3" fill="#b79869"/>
<path d="M429 347H552M429 385H552" stroke="#ba9872" stroke-width="6"/>
<rect x="436" y="236" width="180" height="37" rx="9" fill="#fffdf3"/><text x="526" y="263" text-anchor="middle" fill="#557b98" font-family="Georgia,serif" font-weight="bold" font-size="29" letter-spacing="1">lilbulle</text>
<path d="M415 273H635L649 301H402Z" fill="#a9c2d0"/><path d="M421 273h21l-4 28h-24Zm43 0h21v28h-24Zm43 0h22l3 28h-25Zm43 0h22l5 28h-24Zm44 0h22l10 28h-24Z" fill="#fff9e9"/>
<path d="M402 301v8q8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0 8 13 17 0v-8Z" fill="#a9c2d0"/>
<path d="M408 422H644L662 438H396Z" fill="#f1e7d0"/><path d="M396 438v8h266v-8" fill="#cdbb9a"/>
<path d="M435 396h112v14H435Z" fill="#b6c4a0"/><use href="#flower" transform="translate(446 392)"/><use href="#flower" transform="translate(470 394)"/><use href="#flower" transform="translate(500 391)"/><use href="#flower" transform="translate(530 394)"/>
</g>
<!-- Sale counter and a small sign. -->
<path d="M586 429l71-10 28 16-75 12Z" fill="#e4cda8"/><path d="M610 447v37l75-12v-37Z" fill="#d2b78d"/><path d="M586 429v39l24 16v-37Z" fill="#bea17a"/><path d="M638 426v-16l16-2v16Z" fill="#7c999b"/><path d="M635 410l-4-17 25-3 4 17Z" fill="#bdd5d6"/>
<path d="M379 390l-7 53 40 4-4-51Z" fill="#e1c9a3"/><path d="M379 393l-3 37 31 3-3-35Z" fill="#76948d"/><text x="391" y="412" text-anchor="middle" fill="#fff8e7" font-family="Georgia" font-size="8">OPEN</text><path d="M385 419h12" stroke="#e6edce" stroke-width="2"/>
<!-- Little market booth. -->
<g filter="url(#soft-shadow)"><path d="M836 434l98-19 46 26-98 22Z" fill="#f7e7cb"/><path d="M846 443v74M965 445v74M882 464v78" stroke="#ba9d78" stroke-width="6"/><path d="M831 433l52-40 100 24-50 40Z" fill="#e5bca6"/><path d="M831 433l52 24 50 0-50-23Z" fill="#f9edda"/><path d="M855 419l52 24m-28-40 52 24" stroke="#fff0df" stroke-width="14"/><path d="M848 499l85-17 35 20-85 18Z" fill="#efe0c1"/><path d="M883 520v28l85-19v-27Z" fill="#d0b18c"/><path d="M848 499v29l35 20v-28Z" fill="#e5c9a3"/><text x="908" y="529" text-anchor="middle" fill="#fff9ed" font-family="Georgia" font-size="12" transform="rotate(-12 908 529)">little market</text></g>
<use href="#box" transform="translate(1003 535) scale(.75)"/>
<!-- A tiny delivery cart. -->
<g class="town-cart"><ellipse cx="691" cy="551" rx="45" ry="11" fill="#82916e24"/><path d="M659 522l40-8 29 13-39 10Z" fill="#b7cdd0"/><path d="M659 522v23l30 15v-23Z" fill="#92b4b9"/><path d="M689 537v23l39-8v-25Z" fill="#789d9f"/><path d="M728 529l7-23 15-3" stroke="#7d9388" stroke-width="4" fill="none" stroke-linecap="round"/><circle cx="674" cy="551" r="7" fill="#627d76"/><circle cx="716" cy="552" r="7" fill="#627d76"/><use href="#box" transform="translate(691 515) scale(.7)"/></g>
<!-- Street details. -->
<use href="#lamp" transform="translate(367 458)"/><use href="#lamp" transform="translate(760 542)"/><use href="#lamp" transform="translate(1002 389) scale(.8)"/>
<use href="#mini-tree" transform="translate(146 474) scale(1.05)"/><use href="#mini-tree" transform="translate(251 532) scale(.8)"/><use href="#mini-tree" transform="translate(1051 328)"/><use href="#mini-tree" transform="translate(1080 555) scale(1.1)"/><use href="#mini-tree" transform="translate(412 177) scale(.7)"/>
<use href="#shrub" transform="translate(1080 366)"/><use href="#shrub" transform="translate(794 590)"/><use href="#shrub" transform="translate(815 603) scale(.8)"/><use href="#shrub" transform="translate(369 203)"/>
<use href="#flower" transform="translate(299 564)"/><use href="#flower" transform="translate(317 568)"/><use href="#flower" transform="translate(1038 582)"/>
<path d="M333 493l7-14 34 4-8 14m-24-1 7-14m6 15 7-14" stroke="#fffdf1" stroke-width="6" fill="none"/>
<g fill="#b7c7a0"><ellipse cx="448" cy="549" rx="7" ry="3"/><ellipse cx="473" cy="558" rx="4" ry="2"/><ellipse cx="789" cy="146" rx="6" ry="3"/><ellipse cx="268" cy="194" rx="5" ry="2"/></g>
<text x="468" y="612" font-size="13" fill="#8e9f87" font-family="Georgia" letter-spacing="5" transform="rotate(10 468 612)">A LITTLE CARE, EVERY DAY</text>
</svg>`;
