<script lang="ts">
    import { lang } from "../Scripts/Language";
import DeleteButton from "./DeleteButton.svelte";

    const {
        sourceFont,
    }: {
        sourceFont: Word.Font;
    } = $props();
</script>

<label class="flex hcenter gap">
    {lang("Color")}: <input type="color" value={sourceFont.color || "#000000"} onchange={(e) => (sourceFont.color = e.currentTarget.value)} />
    <DeleteButton
        callback={(e) => {
            // @ts-ignore
            delete sourceFont.color;
        const possibleItem = (e.target as HTMLElement).closest("label")?.querySelector("input[type=color]");
        if (possibleItem) (possibleItem as HTMLInputElement).value = "#000000";
        }}
    ></DeleteButton>
</label><br />
<div class="flex hcenter gap" style="flex-wrap: wrap;">
    {#each [["bold", "Bold"], ["italic", "Italic"], ["strikeThrough", "Strikethrough"], ["doubleStrikeThrough", "Double Strikethrough"], ["superscript", "Superscript"], ["subscript", "Subscript"]] as [fontStyle, fontText]}
        <label class="flex hcenter gap" style="gap: 5px">
            <input
                style="margin-right: 5px"
                type="checkbox"
                bind:checked={sourceFont[fontStyle as "bold"]}
            />
            {lang(fontText)}
        </label>
    {/each}
</div><br />
<label class="flex hcenter gap">
    {lang("Underline")}:
    <div class="selectContainer">
        <select bind:value={sourceFont.underline}>
            {#each ["Mixed", "None", "Hidden", "DotLine", "Single", "Word", "Double", "Thick", "Dotted", "DottedHeavy", "DashLine", "DashLineHeavy", "DashLineLong", "DashLineLongHeavy", "DotDashLine", "DotDashLineHeavy", "TwoDotDashLine", "TwoDotDashLineHeavy", "Wave", "WaveHeavy", "WaveDouble"] as item}
                <option value={item}>{item}</option>
            {/each}
        </select>
    </div>
</label><br />
<label class="flex hcenter gap">
    {lang("Size (in points)")}: <input type="number" bind:value={sourceFont.size} />
</label><br />
<label class="flex hcenter gap">
    {lang("Font name")}: <input type="text" bind:value={sourceFont.name} />
</label><br />
<label class="flex hcenter gap">
    {lang("Highlight color")}: <input
        type="color"
        onchange={(e) => (sourceFont.highlightColor = e.currentTarget.value)}
        value={sourceFont.highlightColor || "#ffffff"}
    />
    <DeleteButton
        callback={(e) => {
            // @ts-ignore
            delete sourceFont.highlightColor;
            const possibleItem = (e.target as HTMLElement).closest("label")?.querySelector("input[type=color]");
            if (possibleItem) (possibleItem as HTMLInputElement).value = "#ffffff";
        }}
    ></DeleteButton>
</label>
