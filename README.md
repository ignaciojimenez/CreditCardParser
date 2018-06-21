# Credit Card Parser
The goal of this simple script is to scan files supossedly containing credit cards against the luhn checksum algorithm and also to well-known credit card formats using regular expresions, contained within a CSV file, and printing everything to stdout.

**Credit Card regex file example (regexcard.csv):**
* Mastercard,\b(5[1-5]\d{2}|222[1-9]|22[3-9]\d|2[3-6]\d{2}|27[01]\d|2720)\d{12}\b
* Maestro,\b(50|5[6-8])\d{14}\|((600[0-9]|6010)\d{2}|(6011(1\d|[56]\d|7[0-3]|7[56]|8[0-5]))|(60(1[2-8]\d|19[0-8]|199|[2-9]\d{2})\d)|61\d{4})\d{10}|((62([01]\d{3}|20\d{2}|21[01]\d|212[0-5]))|62(292[6-9]|29[3-9]\d|3\d{3}|7\d{3}|80\d{2}|81[0-8]\d|819\d))\d{10}|(64[0-3]\d{3}|65060[1-9]|65060[1-9]|6[6-9]\d{4})\d{10}\b
* Visa,\b(4\d{15})\b
* Amex,\b(3[47])\d{13}\b
* JCB,\b(352[89]|35[3-8]\d)\d{12}\b

**Output Example (>>stdout):**
* file1, , , , Visa, , , , , , , , ,
* file2, , , Visa, , , , , , , , ,
* file3, Mastercard, , , , , , , , , , ,
* file4, Mastercard, , , , , , , , , , ,
* file5, Mastercard, , , , , , , , , , ,
* file6, Mastercard, , , , , , , , , , ,
