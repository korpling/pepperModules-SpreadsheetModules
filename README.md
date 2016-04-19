![SaltNPepper project](./gh-site/img/SaltNPepper_logo2010.png)
# pepperModules-SpreadsheetModules
This project provides an importer to support the [Excel](https://products.office.com/de-de/excel) format in linguistic converter framework Pepper (see https://u.hu-berlin.de/saltnpepper). A detailed description of the importer can be found in section [SpreadsheetImporter](#importer).

Pepper is a pluggable framework to convert a variety of linguistic formats (like [TigerXML](http://www.ims.uni-stuttgart.de/forschung/ressourcen/werkzeuge/TIGERSearch/doc/html/TigerXML.html), the [EXMARaLDA format](http://www.exmaralda.org/), [PAULA](http://www.sfb632.uni-potsdam.de/paula.html) etc.) into each other. Furthermore Pepper uses Salt (see https://github.com/korpling/salt), the graph-based meta model for linguistic data, which acts as an intermediate model to reduce the number of mappings to be implemented. That means converting data from a format _A_ to format _B_ consists of two steps. First the data is mapped from format _A_ to Salt and second from Salt to format _B_. This detour reduces the number of Pepper modules from _n<sup>2</sup>-n_ (in the case of a direct mapping) to _2n_ to handle a number of n formats.

![n:n mappings via SaltNPepper](./gh-site/img/puzzle.png)

In Pepper there are three different types of modules:
* importers (to map a format _A_ to a Salt model)
* manipulators (to map a Salt model to a Salt model, e.g. to add additional annotations, to rename things to merge data etc.)
* exporters (to map a Salt model to a format _B_).

For a simple Pepper workflow you need at least one importer and one exporter.

## Requirements
Since the here provided module is a plugin for Pepper, you need an instance of the Pepper framework. If you do not already have a running Pepper instance, click on the link below and download the latest stable version (not a SNAPSHOT):

> Note:
> Pepper is a Java based program, therefore you need to have at least Java 7 (JRE or JDK) on your system. You can download Java from https://www.oracle.com/java/index.html or http://openjdk.java.net/ .


## Install module
If this Pepper module is not yet contained in your Pepper distribution, you can easily install it. Just open a command line and enter one of the following program calls:

**Windows**
```
pepperStart.bat 
```

**Linux/Unix**
```
bash pepperStart.sh 
```

Then type in command *is* and the path from where to install the module:
```
pepper> update org.corpus-tools::pepperModules-pepperModules-EXMARaLDAModules::https://korpling.german.hu-berlin.de/maven2/
```

## Usage
To use this module in your Pepper workflow, put the following lines into the workflow description file. Note the fixed order of xml elements in the workflow description file: &lt;importer/>, &lt;manipulator/>, &lt;exporter/>. The SpreadsheetImporter is an importer module, which can be addressed by one of the following alternatives.
A detailed description of the Pepper workflow can be found on the [Pepper project site](https://u.hu-berlin.de/saltnpepper). 

### a) Identify the module by name

```xml
<importer name="SpreadsheetImporter" path="PATH_TO_CORPUS"/>
```

### b) Identify the module by formats
```xml
<importer formatName="xls" formatVersion="1.0" path="PATH_TO_CORPUS"/>
```
or
```xml
<importer formatName="xlsx" formatVersion="1.0" path="PATH_TO_CORPUS"/>
```

### c) Use properties
```xml
<importer name="SpreadsheetImporter" path="PATH_TO_CORPUS">
  <property key="PROPERTY_NAME">PROPERTY_VALUE</property>
</importer>
```

# <a name="importer">SpreadsheetImporter</a>
At this stage, we want to explain the mapping of a Spreadsheet model to a Salt model. Since there are some conceptual differences between both models, we need to bridge the Spreadsheet model to the graph based Salt model.

## Corpus information and meta data
While you can use a lot of different sheets in e.g. an excel file, by default the first sheet of such a file will be interpreted as the sheet that holds the corpus information, whereas the second sheet will be interpreted as the sheet, that holds the meta data of the given document. You can change those settings by the properties 'corpusSheet', 'metaSheet' and 'metaAnnotation' (See [Properties](#properties) for further information).

## Primary text, tokenization and the timeline  
Since in Salt the anchor of all higher structures and annotations are tokens, we need to identify which column of a corpus sheet represents the primary text and it's tokenization. This is handled by the property "primText" (See the [Properties](#properties) for further information). Each column with a name matching to the value you give to "primText" is used to create a primary text in Salt, by default (if you don't use the property "primText") a column named "tok" will be interpreted as the primary text. Imagine the following Spreadsheet data:

<table>
	<tr><td>tok</td><td>anno1</td><td>anno2</td></tr>
	<tr><td>This</td><td>a11</td><td>a21</td></tr>
	<tr><td>is</td><td>a12</td><td>a22</td></tr>
	<tr><td>a</td><td>a13</td><td>a23</td></tr>
	<tr><td>sample</td><td>a14</td><td>a24</td></tr>
	<tr><td>text</td><td>a15</td><td>a25</td></tr>
	<tr><td>.</td><td>a16</td><td>a26</td></tr>
</table>


Without additional property settings in this sample, there is one primary text "This is a sample text ." since the default primary text column is named "tok". 
Now for each cell in such a column, interpreted as an annotation tier, a token in Salt is created. That means for our sample, the Salt model contains exactly 6 tokens.

Furthermore imagine the following Spreadsheet data:

<table>
	<tr><td>prim1</td><td>primNorm</td><td>anno1</td></tr>
	<tr><td>This</td><td>This</td><td>a1</td></tr>
	<tr><td>is</td><td>is</td><td>a2</td></tr>
	<tr><td>an</td><td>an</td><td>a3</td></tr>
	<tr><td>ex-</td><td rowspan="2">example</td><td>a4</td></tr>
	<tr><td>sample</td><td>a5</td></tr>
	<tr><td>.</td><td>.</td><td>a6</td></tr>
</table>

with the property "primText" set to "prim1, primNorm" this sample contains two primary texts "This is an ex- ample ." and "This is an example .".
Now for each cell in both annotation tiers, a token in Salt is created. That means for this sample, the Salt model contains exactly 11 tokens: 6 tokens connected to the first primary text "prim1" and 5 tokens connected to the second primary text "primNorm". To bring the tokens of both primary texts into a relation, the line number of the Spreadsheet model is mapped to a timeline in the Salt model.


## Annotations

Along with tokens annotations are also modeled as columns in a spreadsheet as shown in the following sample:

<table>
	<tr><td>prim1</td><td>primNorm</td><td>anno1</td><td>anno2</td></tr>
	<tr><td>This</td><td>This</td><td>a11</td><td>a21</td></tr>
	<tr><td>is</td><td>is</td><td>a12</td><td>a22</td></tr>
	<tr><td>an</td><td>an</td><td>a13</td><td>a23</td></tr>
	<tr><td>ex-</td><td rowspan="2">example</td></td><td>a14</td><td rowspan="2">a24</td></tr>
	<<tr><td>sample</td><td>a15</td></tr>
	<tr><td>.</td><td>.</td><td>a16</td><td>a25</td></tr>
</table>

whereas 'anno1' and 'anno2' are annotation tiers and 'prim1' and 'primNorm' are primary text tiers. If your corpus contains more than one primary text tier as in the sample, it's not clear to which primary text a given annotation is related to. Thus you need to specify for each annotation tier to which primary text tier it relates to. This can be managed in the annotation itself, by writing the primary text tier in square brackets behind the annotation name as in the following samle:

<table>
	<tr><td>prim1</td><td>primNorm</td><td>anno1[prim1]</td><td>anno2[primNorm]</td></tr>
	<tr><td>This</td><td>This</td><td>a11</td><td>a21</td></tr>
	<tr><td>is</td><td>is</td><td>a12</td><td>a22</td></tr>
	<tr><td>an</td><td>an</td><td>a13</td><td>a23</td></tr>
	<tr><td>ex-</td><td rowspan="2">example</td><td>a14</td><td rowspan="2">a24</td></tr>
	<tr><td>sample</td><td>a15</td></tr>
	<tr><td>.</td><td>.</td><td>a16</td><td>a25</td></tr>
</table>

If your annotations do not contain those specifications you can add them by the property 'annoPrimRel' without changing your original files, see [Properties](#properties) for further information. Please note that each annotation tier, that is not related to a primary text, will be ignored in the convertion process.

## Meta Annotations
The module currently supports meta annotations of a document only in a specific way. It is assumed that the first column of the sheet, that holds the meta data, contains the meta annotation names, while the second column holds the respective meta annotation value. All other columns will be ignored by the module.

## <a name="properties">Properties</a>
The table contains an overview of all usable properties to customize the behavior of this Pepper module. 

|Name of property	 |Type of property	                                            |optional/ mandatory |	default value |
|--------------------|--------------------------------------------------------------|--------------------|----------------|
|corpusSheet	     |String											             |optional            |	[first sheet] |
|primText			 |primaryTier1, primaryTier2, ...    					         |optional            |	tok |
|metaSheet			 |String			                                             |optional            |	[second sheet] |
|annoPrimRel         |anno1>anno1[primaryTier1], anno2>anno2[primaryTier1], ...		 |optional            |	null |
|setLayer			 |categoryName>{tier1, tier2, tier3}, categoryName2>{tier4}, ... |optional            |	null |
|metaAnnotation		 |Boolean	                                                     |optional            |	true |


### corpusSheet
With the property corpusSheet you can define the sheet that holds the actual corpus information. If you do not set this property, the first sheet will allways be interpreted as the sheet that holds the primary text.

### primText
With the property primText you can define the name of the column(s), that hold the primary text, this can either be a single column name:
```
primText=”TIER_NAME”
```
, or a comma seperated enumeration of column names:
```
primText=”TIER_NAME1, TIER_NAME2, ...”
```
Please make sure that the given tier names are represented in the corpus. If you do not use this property, a column named 'tok' will be considered as the one that holds the primary text.


### metaSheet
With the property metaSheet, you can define the sheet that holds the meta information of the document. If you don not set this property, the second sheet will allways be interpreted as the sheet that holds the meta data of the document. Please note that the first column of this sheet shall hold the meta annotation names, while the second column shall contain the respective meta annotation values. If the sheet contains a meta annotation without a respective value, a warning message will be printed.
```
metaSheet=”SHEET_NAME”
```

### annoPrimRel
In multi-level corpora you need to specify which annotation tier refers to which primary text tier. Therefor the annotation tier name is either followed by the name of the tier, that holds the primary text in square brackets in the annotation itself (in this case you don't need this property), or you set this specification with the property annoPrimRel. A possible key-value set could be:
```
annoPrimRel=”ANNOTATION_NAME>ANNOTATION_NAME[PRIMARY_TEXT1],ANNOTATION_NAME2>ANNOTATION_NAME2[PRIMARY_TEXT2],...”
```


### setLayer
Sometimes it is desirable to add linguistical categories to your corpus, e.g. to get a better overview, for this purpose you can use the property setLayer.
Imagine the following sample:
<table>
	<tr><td>prim1</td><td>primNorm</td><td>anno1</td><td>anno2</td></tr>
	<tr><td>This</td><td>This</td><td>a11</td><td>a21</td></tr>
	<tr><td>is</td><td>is</td><td>a12</td><td>a22</td></tr>
	<tr><td>an</td><td>an</td><td>a13</td><td>a23</td></tr>
	<tr><td>ex-</td><td rowspan="2">example</td><td>a14</td><td rowspan="2">a24</td></tr>
	<<tr><td>sample</td><td>a15</td></tr>
	<tr><td>.</td><td>.</td><td>a16</td><td>a25</td></tr>
</table>

after mapping with the properties: 
```
primText=”prim1, primNorm” annoPrimRel=”anno1>anno1[prim1], anno2>anno2[primNorm]” setLayer=”transcription>{prim1, primNorm}, morphology{anno1, anno2}” 
```
prim1 and primNorm will interpreted as primary texts, whereas anno1 is an annotation of prim1 and anno2 is an annotation of primNorm.
Here we grouped the annotations of the tiers prim1 and primNorm to one SLayer object named transcription and we grouped the annotations of the tiers anno1 and anno2 to another SLayer object named morphology.

### metaAnnotation
If you don't have any meta annotations for your documents, or they don't match the structure needed by the module, you can disable the search for meta annotations by using the property metaAnnotation:
```
metaAnnotation=false
```

## Contribute
Since this Pepper module is under a free license, please feel free to fork it from github and improve the module. If you even think that others can benefit from your improvements, don't hesitate to make a pull request, so that your changes can be merged.
If you have found any bugs, or have some feature request, please open an issue on github. If you need any help, please write an e-mail to saltnpepper@lists.hu-berlin.de .

## Funders
This project has been funded by the [department of corpus linguistics and morphology](https://www.linguistik.hu-berlin.de/institut/professuren/korpuslinguistik/) of the Humboldt-Universität zu Berlin, the Institut national de recherche en informatique et en automatique ([INRIA](www.inria.fr/en/)) and the [Sonderforschungsbereich 632](https://www.sfb632.uni-potsdam.de/en/). 

## License
  Copyright 2009 Humboldt-Universität zu Berlin, INRIA.

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at
 
  http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
