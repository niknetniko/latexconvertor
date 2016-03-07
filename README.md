# latexconvertor
Convert a PowerPoint with speaker's notes to a LaTeX file. This is really simular to the "print with speaker's notes" in Office.

## Features &limitations

Note: this is alpha software. Check the result before using/distributing the result.

Following features are currently supported:

* Recognizes paragraphs.
* Should recognize all Unicode characters.
* Recognizes sub- and superscript.
* Recognizes lists

Following things do currently not work or do not work as expected:

* All text that is indented will be recognized as a list.
* All lists are recognized as unordered list.
* Does not recognize bold/italic/underline.
* Nested lists are untested.
* The interface is in Dutch (feature, not a bug)

## LaTeX

* The file outputted should be valid LaTeX.
* The LaTeX is set up to use A4-wide, Latin Modern, and UTF-8 support (without BOM).

## Installation

1. Download the .ppam plugin file.
2. Open the PowerPoint.
3. Go to the Developers tab (in the Ribbon).
4. Go to plugins, and select the downloaded file. Click OK on any warnings.
5. Go to the new Plugins tab (in the Ribbon).
6. Press the button.
7. Sit back.
8. Profit?


## License

```
Copyright 2016 Niko Strijbol

In licentie gegeven krachtens de EUPL, versie 1.1 of –
zodra deze worden goedgekeurd door de Europese Commissie
- opeenvolgende versies van de EUPL (de "licentie");
U mag dit werk niet gebruiken, behalve onder de
voorwaarden van de licentie.
U kunt een kopie van de licentie vinden op:
https://joinup.ec.europa.eu/software/page/eupl5

Tenzij dit op grond van toepasselijk recht vereist is
of schriftelijk is overeengekomen, wordt software
krachtens deze licentie verspreid "zoals deze is",
ZONDER ENIGE GARANTIES OF VOORWAARDEN,
noch expliciet noch impliciet.
Zie de licentie voor de specifieke bepalingen voor
toestemmingen en beperkingen op grond van de licentie.



Licensed under the EUPL, Version 1.1 or – as soon they
will be approved by the European Commission - subsequent
versions of the EUPL (the "Licence");
You may not use this work except in compliance with the
Licence.
You may obtain a copy of the Licence at:
https://joinup.ec.europa.eu/software/page/eupl5

Unless required by applicable law or agreed to in
writing, software distributed under the Licence is
distributed on an "AS IS" basis,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either
express or implied.
See the Licence for the specific language governing
permissions and limitations under the Licence.
```
