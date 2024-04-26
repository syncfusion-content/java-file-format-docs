---
title: Working with Mathematical Equation | Syncfusion
description: This section illustrates about create, modify and remove mathematical equation in Word document without MS Word or Office interop
platform: java-file-formats
control: Word Library
documentation: UG
---
# Working with Mathematical Equation

Equations in Word document are combination of mathematical symbols or text. For example, you can create a Fourier series equation in Word document.

![Mathematical equation in Microsoft Word document](WorkingwithMathematicalEquation_images/Mathematical Equation.png)

## Types of equation

The following different structures of equation can be created by using the Essential DocIO.

![Different structures of equation in Microsoft Word application](WorkingwithMathematicalEquation_images/EquationStructures.png)

* Accent
* Bar
* Box
* Border box
* Delimiter
* Equation array
* Fraction
* Function
* Group character
* Limit
* Matrix
* N-Array
* Radical
* Phantom
* SubSuperscript
* Left SubSuperscript
* Right SubSuperscript

### Accent

You can add an accent mark to the equation. The following code example shows how to add an accent mark to the equation.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation  to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds an accent equation
IOfficeMathAccent mathAccent = (IOfficeMathAccent) officeMath.getFunctions().add(MathFunctionType.Accent);
//Sets the accent character
mathAccent.setAccentCharacter("̆");
//Adds the run element for accent
IOfficeMathRunElement officeMathRunElement = (IOfficeMathRunElement) mathAccent.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
WTextRange textRange =(WTextRange) officeMathRunElement.getItem();
//Sets text for accent equation
textRange.setText("a");
//Applies character formatting for text range
textRange.getCharacterFormat().setBold(true);
textRange.getCharacterFormat().setItalic(true);
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Bar

You can add a bar (which adds horizontal line on top or bottom) to the equation. The following code example shows how to add a bar to the equation.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Add a section and a paragraph in the empty document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a bar equation
IOfficeMathBar mathBar = (IOfficeMathBar) officeMath.getFunctions().add(0, MathFunctionType.Bar);
//Sets the position of bar
mathBar.setBarTop(true);
//Adds the run element for bar
IOfficeMathRunElement officeMathRunElement = (IOfficeMathRunElement) mathBar.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for bar equation
((WTextRange)officeMathRunElement.getItem()).setText("a");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Box

You can add a box to the equation. The following code example shows how to add a box to the equation.

{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a box equation
IOfficeMathBox mathBox = (IOfficeMathBox) officeMath.getFunctions().add(0, MathFunctionType.Box);
//Adds the run element for box
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMath.getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for math
((WTextRange)officeMathRunElement.getItem()).setText("a+b");
//Enables the flag, to behave the box and its contents as a single operator
mathBox.setOperatorEmulator(true);
//Enables the flag, to act box as the mathematical differential
mathBox.setEnableDifferential(true);
//Adds a break in box equation
mathBox.setBreak(officeMath.getBreaks().add(0));
//Adds the run element for box
officeMathRunElement =(IOfficeMathRunElement) mathBox.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
//Sets text for box equation
((WTextRange)officeMathRunElement.getItem()).setText( "==");
//Adds the run element for box
officeMathRunElement =(IOfficeMathRunElement) mathBox.getEquation().getFunctions().add(1, MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
//Sets text for box equation
((WTextRange)officeMathRunElement.getItem()).setText("adx");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Border box

You can add a box with the borders on four sides and strikethrough on horizontal, vertical, and diagonal directions to the equation. The following code example shows how to add a border box to the equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a border box equation
IOfficeMathBorderBox mathBorderBox =(IOfficeMathBorderBox) officeMath.getFunctions().add(0, MathFunctionType.BorderBox);
//Sets the diagonal strikethrough from lower left to upper right
mathBorderBox.setStrikeDiagonalUp(true);
//Sets the diagonal strikethrough from upper left to lower right
mathBorderBox.setStrikeDiagonalDown(true);
//Sets the horizontal strikethrough
mathBorderBox.setStrikeHorizontal(true);
//Sets the vertical strikethrough
mathBorderBox.setStrikeVertical(true);
//Enables the flag, to hide the bottom border of an equation
mathBorderBox.setHideBottom(true);
//Enables the flag, to hide the left border of an equation
mathBorderBox.setHideLeft(true);
//Sets false to show the right border of an equation
mathBorderBox.setHideRight(false);
//Sets false to show the top border of an equation
mathBorderBox.setHideTop(false);
//Adds the run element for border box
IOfficeMathRunElement officeMathRunElement = (IOfficeMathRunElement) mathBorderBox.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for border box equation
((WTextRange)officeMathRunElement.getItem()).setText("a+b-c");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();

{% endhighlight %}

{% endtabs %}

### Delimiter

You can add a delimiter (parenthesis, square brackets and other characters) to the equation. The following code example shows how to a add delimiter to the equation. 
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a delimiter equation
IOfficeMathDelimiter mathDelimiter =(IOfficeMathDelimiter) officeMath.getFunctions().add(0, MathFunctionType.Delimiter);
//Sets the begin character
mathDelimiter.setBeginCharacter("[");
//Sets the end character
mathDelimiter.setEndCharacter("]");
//Enables the flag, to grow delimiter characters to full height of the arguments
mathDelimiter.setIsGrow(true);
//Sets the appearance of delimiters
mathDelimiter.setDelimiterShape(MathDelimiterShapeType.Match);
//Adds the run element for delimiter
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) mathDelimiter.getEquation().add(0).getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for delimiter equation
((WTextRange)officeMathRunElement.getItem()).setText("a+b");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Equation array

You can create a one dimensional array of equations in Word document. The following code example shows how to create an array of equations.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds an equation array
IOfficeMathEquationArray mathEquationArray =(IOfficeMathEquationArray) officeMath.getFunctions().add(0, MathFunctionType.EquationArray);
//Sets the vertical alignment for equation array
mathEquationArray.setVerticalAlignment(MathVerticalAlignment.Center);
//Enables the flag, to distribute the equation array equally within the container
mathEquationArray.setExpandEquationContainer(true);
//Enables the flag, to expand the equations in an equation array to the maximum width
mathEquationArray.setExpandEquationContainer(true);
//Sets the row spacing rule
mathEquationArray.setRowSpacingRule(SpacingRule.Multiple);
//Adds the run element for equation array
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) mathEquationArray.getEquation().add(0).getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for equation) array
((WTextRange)officeMathRunElement.getItem()).setText("x+y+z=0");
//Adds the run element for equation array
officeMathRunElement =(IOfficeMathRunElement) mathEquationArray.getEquation().add(1).getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for equation array
((WTextRange)officeMathRunElement.getItem()).setText("x+y-z=1");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Fraction

You can create a fraction equation with a numerator and denominator in Word document. The following code example shows how to create a fraction equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a fraction getEquation()
IOfficeMathFraction mathFraction =(IOfficeMathFraction) officeMath.getFunctions().add(0, MathFunctionType.Fraction);
//Sets the denominator for fraction
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) mathFraction.getNumerator().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("a");
//Sets the Numerator for fraction
officeMathRunElement =(IOfficeMathRunElement) mathFraction.getDenominator().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("b");
//Sets the fraction type
mathFraction.setFractionType(MathFractionType.NormalFractionBar);
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Function

You can create trigonometric functions in a Word document. The following code example shows how to create a function.  
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a function
IOfficeMathFunction mathFunction =(IOfficeMathFunction) officeMath.getFunctions().add(0, MathFunctionType.Function);
//Sets the function name
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) mathFunction.getFunctionName().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("sin");
//Adds the run element for function
officeMathRunElement =(IOfficeMathRunElement) mathFunction.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for function
((WTextRange)officeMathRunElement.getItem()).setText("90");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Group character

You can group mathematical equations by adding a grouping character at above or below to the corresponding equations. The following code example shows how to create an equation with grouping character.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath math = document.getLastParagraph().appendMath();
//Adds a new math
IOfficeMath officeMath = math.getMathParagraph().getMaths().add();
//Adds a group character equation
IOfficeMathGroupCharacter officeMathGroupCharacter =(IOfficeMathGroupCharacter) officeMath.getFunctions().add(0, MathFunctionType.GroupCharacter);
//Sets the group character
officeMathGroupCharacter.setGroupCharacter("⏞");
//Enables the flag to align group character at top
officeMathGroupCharacter.setHasAlignTop(true);
//Enables the flag to align the text and group character
officeMathGroupCharacter.setHasCharacterTop(true);
//Adds the run element for group character
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathGroupCharacter.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for group character equation
((WTextRange)officeMathRunElement.getItem()).setText("a-b");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Limit

You can add upper limit or lower limit to the mathematical equation. The following code example shows how to create limit equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wMath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wMath.getMathParagraph().getMaths().add();
//Adds function to the math
IOfficeMathFunction officeMathFunction =(IOfficeMathFunction) officeMath.getFunctions().add(0, MathFunctionType.Function);
//Adds a mathematical limit equation
IOfficeMathLimit officeMathLimit =(IOfficeMathLimit) officeMathFunction.getFunctionName().getFunctions().add(0, MathFunctionType.Limit);
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathLimit.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
//Sets text for limit equation.
((WTextRange)officeMathRunElement.getItem()).setText("lim");
//Sets the type of the limit.
officeMathLimit.setLimitType(MathLimitType.LowerLimit);
IOfficeMathRunElement officeMathRunElement_limit =(IOfficeMathRunElement) officeMathLimit.getLimit().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement_limit.setItem(new WTextRange(document));
//Sets the Limit value.
((WTextRange)officeMathRunElement_limit.getItem()).setText("n=0");
officeMathLimit.setLimitType( MathLimitType.LowerLimit);
officeMathRunElement =(IOfficeMathRunElement) officeMathFunction.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
//Sets text for base of the specified equation
((WTextRange)officeMathRunElement.getItem()).setText("x");
//Saves the Word document
document.save("Sample.docx");
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Matrix

You can create a matrix equation in a Word document. The following code example shows how to create a matrix equation. 
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds matrix getEquation()
IOfficeMathMatrix mathMatrix = (IOfficeMathMatrix) officeMath.getFunctions().add(MathFunctionType.Matrix);
//Sets vertical alignment for matrix
mathMatrix.setVerticalAlignment(MathVerticalAlignment.Center);
//Sets width for matrix columns
mathMatrix.setColumnWidth(1);
//Sets column spacing rule
mathMatrix.setColumnSpacingRule(SpacingRule.OneAndHalf);
//Sets column spacing value
mathMatrix.setColumnSpacing(3);
//Enables the flag to hide place holders
mathMatrix.setHidePlaceHolders(true);
//Sets row spacing rule.
mathMatrix.setRowSpacingRule(SpacingRule.Double);
//Sets row spacing value.
mathMatrix.setRowSpacing(2);

//Adds a new column
mathMatrix.getColumns().add();
//Adds a new row
mathMatrix.getRows().add();
//Sets horizontal alignment for column
mathMatrix.getColumns().get(0).setHorizontalAlignment(MathHorizontalAlignment.Left);

//Gets an argument in first cell in first row
officeMath = mathMatrix.getRows().get(0).getArguments().get(0);
//Sets text for argument in first cell in first row
IOfficeMathRunElement officeMathRunElement = (IOfficeMathRunElement) officeMath.getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("1");

//Adds a new column
mathMatrix.getColumns().add();
//Adds a new row
mathMatrix.getRows().add();
//Gets an argument in second cell in first row
officeMath = mathMatrix.getRows().get(0).getArguments().get(1);
//Sets text for argument in second cell in first row
officeMathRunElement = (IOfficeMathRunElement) officeMath.getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("2");

//Gets an argument in first cell in second row
officeMath = mathMatrix.getRows().get(1).getArguments().get(0);
//Sets text for argument in first cell in second row
officeMathRunElement = (IOfficeMathRunElement) officeMath.getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("3");

//Gets an argument in second cell in second row
officeMath = mathMatrix.getRows().get(1).getArguments().get(1);
//Sets text for argument in second cell in second row
officeMathRunElement = (IOfficeMathRunElement) officeMath.getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("4");
//Saves the Word document.
document.save("Sample.docx");
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### N-Array

You can create an equation with common large operators such as summation, integrals, union, intersection, logical OR, logical AND, products and co-products. The following code example shows how to create a summation with limits.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wMath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wMath.getMathParagraph().getMaths().add();
//Adds a N-Array equation
IOfficeMathNArray officeMathNArray = (IOfficeMathNArray) officeMath.getFunctions().add(0, MathFunctionType.NArray);
//Sets N-Array character.
officeMathNArray.setNArrayCharacter("∑");
//Enables the flag, to grow N-array character to full height of the arguments
officeMathNArray.setHasGrow(false);
//Enables the flag to hide lower limit
officeMathNArray.setHideLowerLimit(false);
//Enables the flag to hide upper limit
officeMathNArray.setHideUpperLimit(false);
//Sets false to set limit position on above the summation
officeMathNArray.setSubSuperscriptLimit(false);
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathNArray.getSubscript().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for superscript property of NArray equation
((WTextRange)officeMathRunElement.getItem()).setText("n=1");
officeMathRunElement =(IOfficeMathRunElement) officeMathNArray.getSuperscript().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("10");
officeMathRunElement =(IOfficeMathRunElement) officeMathNArray.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for NArray equation
((WTextRange)officeMathRunElement.getItem()).setText("x");
//Saves the Word document.
document.save("Sample.docx");
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Radical

You can create a radical equation in Word document. The following example shows how to create a radical equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds a radical equation
IOfficeMathRadical officeMathRadical = (IOfficeMathRadical) officeMath.getFunctions().add(0, MathFunctionType.Radical);
//Sets false to show degree in radical
officeMathRadical.setHideDegree(false);
//Adds a degree for radical equation
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathRadical.getDegree().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("2");
//Adds an run element for radical
officeMathRunElement =(IOfficeMathRunElement) officeMathRadical.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets the radicand text for radical equation
((WTextRange)officeMathRunElement.getItem()).setText("x");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Phantom

You can create a phantom equation to add the spacing of the phantom
without displaying that base and suppressing part of the glyph from spacing considerations. The following code example shows how to create a phantom equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds a radical equation
IOfficeMathRadical officeMathRadical = (IOfficeMathRadical) officeMath.getFunctions().add(0, MathFunctionType.Radical);
//Adds a degree for radical
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathRadical.getDegree().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
((WTextRange)officeMathRunElement.getItem()).setText("2");
//Adds a phantom equation in radical.
IOfficeMathPhantom officeMathPhantom =(IOfficeMathPhantom) officeMathRadical.getEquation().getFunctions().add(0, MathFunctionType.Phantom);
//Enables the flag, to show the contents of phantom
officeMathPhantom.setShow(true);
//Enables the flag, to transparent the phantom
officeMathPhantom.setTransparent(true);
//Enables the flag, to ignore the ascent of the phantom contents in spacing
officeMathPhantom.setZeroAscent(true);
//Enables the flag, to ignore the descent of the phantom contents in spacing
officeMathPhantom.setZeroDescent(true);
//Enables the flag, to ignore the width of a phantom contents in spacing
officeMathPhantom.setZeroWidth(true);
//Adds a run element for phantom
officeMathRunElement =(IOfficeMathRunElement) officeMathPhantom.getEquation().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for phantom equation
((WTextRange)officeMathRunElement.getItem()).setText("a-b");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### SubSuperscript

You can add a superscript or subscript equation in a Word document. The following code shows how to create a superscript equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds a superscript equation
IOfficeMathScript officeMathScript = (IOfficeMathScript) officeMath.getFunctions().add(0, MathFunctionType.SubSuperscript);
//Sets the type of the script as superscript.
officeMathScript.setScriptType(MathScriptType.Superscript);
//Adds a run element for superscript.
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathScript.getScript().getFunctions().add(MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
WTextRange textRange = (WTextRange) officeMathRunElement.getItem();
//Sets text for superscript.
textRange.setText("2");
//Adds run element for equation
officeMathRunElement =(IOfficeMathRunElement) officeMathScript.getEquation().getFunctions().add(MathFunctionType.RunElement);
//Sets text for equation
((WTextRange)officeMathRunElement.getItem()).setText("x");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Left SubSuperscript

You can add superscript and subscript on the left side of mathematical equation. The following code example shows how to add superscript and subscript on the left side of the equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds a left subsuperscript equation
IOfficeMathLeftScript officeMathLeftSubScript = (IOfficeMathLeftScript) officeMath.getFunctions().add(0, MathFunctionType.LeftSubSuperscript);
//Adds run element for left subscript
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathLeftSubScript.getSubscript().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for subscript
((WTextRange)officeMathRunElement.getItem()).setText("1");
//Adds a run element for left superscript
officeMathRunElement =(IOfficeMathRunElement) officeMathLeftSubScript.getSuperscript().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for left superscript.
((WTextRange)officeMathRunElement.getItem()).setText("n");
officeMathRunElement =(IOfficeMathRunElement) officeMathLeftSubScript.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem( new WTextRange(document));
//Sets text for equation
((WTextRange)officeMathRunElement.getItem()).setText("Y");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

### Right SubSuperscript

You can add superscript and subscript on the right side of mathematical equation. The following code example shows how to add superscript and subscript on the right side of the equation.
{% tabs %}
{% highlight JAVA %}
//Creates a new Word document
WordDocument document = new WordDocument();
//Adds one section and one paragraph to the document
document.ensureMinimal();
//Appends a new mathematical equation to the paragraph
WMath wmath = document.getLastParagraph().appendMath();
IOfficeMath officeMath = wmath.getMathParagraph().getMaths().add();
//Adds a right subsuperscript equation
IOfficeMathRightScript officeMathRightScript = (IOfficeMathRightScript) officeMath.getFunctions().add(0, MathFunctionType.RightSubSuperscript);
//Sets false to align subscript and superscript horizontally
officeMathRightScript.setIsSkipAlign(true);
//Adds run element for right subscript.
IOfficeMathRunElement officeMathRunElement =(IOfficeMathRunElement) officeMathRightScript.getSubscript().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for right subscript
((WTextRange)officeMathRunElement.getItem()).setText("1");
//Adds run element for right superscript
officeMathRunElement =(IOfficeMathRunElement) officeMathRightScript.getSuperscript().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for right superscript.
((WTextRange)officeMathRunElement.getItem()).setText("n");
officeMathRunElement =(IOfficeMathRunElement) officeMathRightScript.getEquation().getFunctions().add(0, MathFunctionType.RunElement);
officeMathRunElement.setItem(new WTextRange(document));
//Sets text for equation
((WTextRange)officeMathRunElement.getItem()).setText("Y");
//Saves the Word document
document.save("Sample.docx", FormatType.Docx);
//Closes the document
document.close();
{% endhighlight %}

{% endtabs %}

## Modify existing equation

You can add or modify the text and formatting of existing mathematical equation in Word document. The following screenshots shows an existing mathematical equation in the input Word document. 
![Existing mathematical equation in Word document](WorkingwithMathematicalEquation_images/EditEquation.png)

The following code example shows how to modify an existing mathematical equation in the Word document.
{% tabs %}
{% highlight JAVA %}
//Opens an existing Word document
WordDocument document = new WordDocument("Template.docx");
//Access the paragraph from Word document
WParagraph paragraph = (WParagraph) document.getLastSection().getBody().getChildEntities().get(0);
//Access the mathematical equation from the paragraph
WMath math = (WMath) paragraph.getChildEntities().get(0);
//Access the radical equation
IOfficeMathRadical mathRadical = (IOfficeMathRadical) math.getMathParagraph().getMaths().get(0).getFunctions().get(1);
//Access the fraction equation in radical
IOfficeMathFraction mathFraction = (IOfficeMathFraction) mathRadical.getEquation().getFunctions().get(0);
//Access the n-array equation in fraction
IOfficeMathNArray mathNAry = (IOfficeMathNArray) mathFraction.getNumerator().getFunctions().get(0);
//Access the math script in n-array
IOfficeMathScript mathScript = (IOfficeMathScript) mathNAry.getEquation().getFunctions().get(0);
//Access the delimiter in math script
IOfficeMathDelimiter mathDelimiter = (IOfficeMathDelimiter) mathScript.getEquation().getFunctions().get(0);
//Removes the delimiter
mathScript.getEquation().getFunctions().remove(mathDelimiter);
//Modifies the run element in math script
IOfficeMathRunElement MathParagraphItem = (IOfficeMathRunElement) mathScript.getEquation().getFunctions().add(MathFunctionType.RunElement);
MathParagraphItem.setItem( new WTextRange(document));
//Sets the text value
((WTextRange)MathParagraphItem.getItem()).setText("x");
//Applies character format to the text
((WTextRange)MathParagraphItem.getItem()).getCharacterFormat().setItalic(true);
((WTextRange)MathParagraphItem.getItem()).getCharacterFormat().setFontSize(20);
//Applies math format to the text
MathParagraphItem.getMathFormat().setStyle(MathStyleType.Italic);
//Saves the word document
document.save("Sample.docx");
//close the word document
document.close();
{% endhighlight %}

{% endtabs %}

By executing the above code example, it generates output Word document as follows.
![Resultant output Word document](WorkingwithMathematicalEquation_images/EditedEquation.png)