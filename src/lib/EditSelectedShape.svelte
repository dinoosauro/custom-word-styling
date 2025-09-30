<script lang="ts">
    import { AvailableShapes } from "../Scripts/HelperType";
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
            {#each AvailableShapes as option}
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
