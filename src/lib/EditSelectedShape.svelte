<script lang="ts">
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    const { shape }: { shape: Word.Shape } = $props();
</script>

<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={shape.allowOverlap} />
    {lang("Allow shape overlapping")}
</label><br />
<label class="flex hcenter gap">
    {lang("Geometric shape type")}:
    <div class="selectContainer">
        <select bind:value={shape.geometricShapeType}>
            {#each ["LineInverse", "Triangle", "RightTriangle", "Rectangle", "Diamond", "Parallelogram", "Trapezoid", "NonIsoscelesTrapezoid", "Pentagon", "Hexagon", "Heptagon", "Octagon", "Decagon", "Dodecagon", "Star4", "Star5", "Star6", "Star7", "Star8", "Star10", "Star12", "Star16", "Star24", "Star32", "RoundRectangle", "Round1Rectangle", "Round2SameRectangle", "Round2DiagonalRectangle", "SnipRoundRectangle", "Snip1Rectangle", "Snip2SameRectangle", "Snip2DiagonalRectangle", "Plaque", "Ellipse", "Teardrop", "HomePlate", "Chevron", "PieWedge", "Pie", "BlockArc", "Donut", "NoSmoking", "RightArrow", "LeftArrow", "UpArrow", "DownArrow", "StripedRightArrow", "NotchedRightArrow", "BentUpArrow", "LeftRightArrow", "UpDownArrow", "LeftUpArrow", "LeftRightUpArrow", "QuadArrow", "LeftArrowCallout", "RightArrowCallout", "UpArrowCallout", "DownArrowCallout", "LeftRightArrowCallout", "UpDownArrowCallout", "QuadArrowCallout", "BentArrow", "UturnArrow", "CircularArrow", "LeftCircularArrow", "LeftRightCircularArrow", "CurvedRightArrow", "CurvedLeftArrow", "CurvedUpArrow", "CurvedDownArrow", "SwooshArrow", "Cube", "Can", "LightningBolt", "Heart", "Sun", "Moon", "SmileyFace", "IrregularSeal1", "IrregularSeal2", "FoldedCorner", "Bevel", "Frame", "HalfFrame", "Corner", "DiagonalStripe", "Chord", "Arc", "LeftBracket", "RightBracket", "LeftBrace", "RightBrace", "BracketPair", "BracePair", "Callout1", "Callout2", "Callout3", "AccentCallout1", "AccentCallout2", "AccentCallout3", "BorderCallout1", "BorderCallout2", "BorderCallout3", "AccentBorderCallout1", "AccentBorderCallout2", "AccentBorderCallout3", "WedgeRectCallout", "WedgeRRectCallout", "WedgeEllipseCallout", "CloudCallout", "Cloud", "Ribbon", "Ribbon2", "EllipseRibbon", "EllipseRibbon2", "LeftRightRibbon", "VerticalScroll", "HorizontalScroll", "Wave", "DoubleWave", "Plus", "FlowChartProcess", "FlowChartDecision", "FlowChartInputOutput", "FlowChartPredefinedProcess", "FlowChartInternalStorage", "FlowChartDocument", "FlowChartMultidocument", "FlowChartTerminator", "FlowChartPreparation", "FlowChartManualInput", "FlowChartManualOperation", "FlowChartConnector", "FlowChartPunchedCard", "FlowChartPunchedTape", "FlowChartSummingJunction", "FlowChartOr", "FlowChartCollate", "FlowChartSort", "FlowChartExtract", "FlowChartMerge", "FlowChartOfflineStorage", "FlowChartOnlineStorage", "FlowChartMagneticTape", "FlowChartMagneticDisk", "FlowChartMagneticDrum", "FlowChartDisplay", "FlowChartDelay", "FlowChartAlternateProcess", "FlowChartOffpageConnector", "ActionButtonBlank", "ActionButtonHome", "ActionButtonHelp", "ActionButtonInformation", "ActionButtonForwardNext", "ActionButtonBackPrevious", "ActionButtonEnd", "ActionButtonBeginning", "ActionButtonReturn", "ActionButtonDocument", "ActionButtonSound", "ActionButtonMovie", "Gear6", "Gear9", "Funnel", "MathPlus", "MathMinus", "MathMultiply", "MathDivide", "MathEqual", "MathNotEqual", "CornerTabs", "SquareTabs", "PlaqueTabs", "ChartX", "ChartStar", "ChartPlus"] as option}
                <option value={option}>{option}</option>
            {/each}
        </select>
    </div>
</label><br />
<Card secondCard={true}>
    <h3>{lang("Shape position")}:</h3>
    <label class="flex hcenter gap">
        {lang("Width")}: <input type="number" bind:value={shape.width} />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Height")}: <input type="number" bind:value={shape.height} />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Top")}: <input type="number" bind:value={shape.top} />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Left")}: <input type="number" bind:value={shape.top} />
    </label>
</Card><br />
<Card secondCard={true}>
    <h3>{lang("Text frame settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Orientation")}:
        <div class="selectContainer">
            <select bind:value={shape.textFrame.orientation}>
                <option value="None">{lang("Standard")}</option>
                <option value="Horizontal">{lang("Horizontal")}</option>
                <option value="EastAsianHorizontalRotated"
                    >{lang("Horizontal (East-Asian) (Rotated)")}</option
                >
                <option value="EastAsianVertical"
                    >{lang("Vertical (East-Asian)")}</option
                >
                <option value="Vertical">{lang("Vertical")}</option>
                <option value="Vertical270">{lang("Vertical (270°)")}</option>
                <option value="Mixed">Mixed</option>
            </select>
        </div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("Text alignment")}:
        <div class="selectContainer">
            <select bind:value={shape.textFrame.verticalAlignment}>
                <option value="Top">{lang("Top")}</option>
                <option value="Middle">{lang("Middle")}</option>
                <option value="Bottom">{lang("Bottom")}</option>
            </select>
        </div>
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={shape.textFrame.wordWrap} />{lang(
            "Enable word wrapping",
        )}
    </label><br />
    <label class="flex hcenter gap">
        <input
            type="checkbox"
            bind:checked={shape.textFrame.noTextRotation}
        />{lang("Disable text rotation when rotating the shape")}
    </label><br />
    <Card>
        <h4>{lang("Margin")}:</h4>
        <label class="flex hcenter gap"
            >{lang("Top margin")}:
            <input
                type="number"
                bind:value={shape.textFrame.topMargin}
            /></label
        ><br />
        <label class="flex hcenter gap"
            >{lang("Left margin")}:
            <input
                type="number"
                bind:value={shape.textFrame.leftMargin}
            /></label
        ><br />
        <label class="flex hcenter gap"
            >{lang("Right margin")}:
            <input
                type="number"
                bind:value={shape.textFrame.rightMargin}
            /></label
        ><br />
        <label class="flex hcenter gap"
            >{lang("Bottom margin")}:
            <input
                type="number"
                bind:value={shape.textFrame.bottomMargin}
            /></label
        >
    </Card>
</Card>
