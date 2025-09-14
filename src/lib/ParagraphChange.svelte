<script lang="ts">
    import type { HelperType } from "../Scripts/HelperType";
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    import HelperDialogs from "./HelperDialogs.svelte";

    const {
        sourceParagraph,
        isParagraphFormat
    }: {
        sourceParagraph: Word.ParagraphFormat | Word.Paragraph;
        isParagraphFormat: boolean
    } = $props();
    let helperProp: HelperType | undefined = $state(undefined);
</script>

<label class="flex hcenter gap">
    {lang("Alignment")}: 
    <div class="selectContainer">
        <select bind:value={sourceParagraph.alignment}>
            {#each ["Unknown", "Left", "Centered", "Right", "Justified"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select>
    </div>
</label><br />
<label class="flex hcenter gap">
    <span class="help" onclick={() => (helperProp = "Indent")}>{lang("Indent (in points)")}</span>
    <input type="number" bind:value={sourceParagraph.firstLineIndent} />
</label><br />
{#if isParagraphFormat}
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={(sourceParagraph as Word.ParagraphFormat).keepTogether} />
        <span class="help" onclick={() => (helperProp = "ParagraphSamePage")}>{lang("Keep all the lines of the paragraph in the same page")}</span>
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={(sourceParagraph as Word.ParagraphFormat).keepWithNext} />
        <span class="help" onclick={() => (helperProp = "ParagraphSamePage")}>{lang("Put this paragraph in the same page as the next paragraph")}</span>
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={(sourceParagraph as Word.ParagraphFormat).widowControl} />
        {lang("Avoid sending the second or last line to a new page (the first or second-last line will be moved too)")}
</label><br />
{/if}
<label class="flex hcenter gap">
    <span class="help" onclick={() => (helperProp = "LevelOutline")}>{lang("Level outline")}:</span>
    <div class="selectContainer">
        <select bind:value={sourceParagraph.outlineLevel}>
            {#each ["1", "2", "3", "4", "5", "6", "7", "8", "9", "BodyText"] as option}
                <option value={`OutlineLevel${option}`}>{lang("Level")} {option}</option>
            {/each}
        </select>
    </div>
</label><br />
<Card secondCard={true}>
    <h3>{lang("Indentation")}:</h3>
    <label class="flex hcenter gap" style="flex-wrap: wrap;">
        <label class="flex hcenter gap" style="gap: 5px">
            {lang("Left")}: <input
                type="number"
                style="width: 60px;"
                bind:value={sourceParagraph.leftIndent}
            />
        </label>
        <label class="flex hcenter gap" style="gap: 5px">
            {lang("Right")}: <input
                type="number"
                style="width: 60px;"
                bind:value={sourceParagraph.rightIndent}
            />
        </label>
    </label>
</Card>
<br />
<Card secondCard={true}>
    <h3>{lang("Spacing")}:</h3>
    <label class="flex hcenter gap">
        {lang("Line spacing")}: <input
            type="number"
            bind:value={sourceParagraph.lineSpacing}
        />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Space before the paragraph")}: <input
            type="number"
            bind:value={sourceParagraph.spaceBefore}
        />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Space after the paragraph")}: <input
            type="number"
            bind:value={sourceParagraph.spaceAfter}
        />
    </label>
</Card>

{#if helperProp}
<HelperDialogs helperType={helperProp} callback={() => (helperProp = undefined)}></HelperDialogs>
{/if}