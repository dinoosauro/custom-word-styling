<script lang="ts">
    import type { HelperType } from "../Scripts/HelperType";
    import { lang } from "../Scripts/Language";
    import HelperDialogs from "./HelperDialogs.svelte";

    const {
        sourceTable,
    }: {
        sourceTable: Word.TableStyle;
    } = $props();
    let helperProp: HelperType | undefined = $state(undefined);
</script>

<label class="flex hcenter gap">
    {lang("Alignment")}:
    <div class="selectContainer">
        <select bind:value={sourceTable.alignment}>
            {#each ["Unknown", "Left", "Centered", "Right", "Justified"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select>
    </div>
</label><br />
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={sourceTable.allowBreakAcrossPage} />
    {lang("The lines can be broken across pages")}
</label><br />
<label class="flex hcenter gap">
    <span class="help" onclick={() => (helperProp = "CellSpacing")}>{lang("Cell spacing")}:</span> <input type="number" bind:value={sourceTable.cellSpacing} />
</label><br />
<div class="secondCard">
    <h3>{lang("Cell margin")}:</h3>
    <label class="flex hcenter gap" style="flex-wrap: wrap">
        {#each [["top", "Top"], ["bottom", "Bottom"], ["left", "Left"], ["right", "Right"]] as [key, title]}
            <label class="flex hcenter gap" style="gap: 5px">
                {lang(title)}
                <input
                    type="number"
                    style="width: 60px"
                    bind:value={sourceTable[`${key as "top"}CellMargin`]}
                />
            </label>
        {/each}
    </label>
</div>

{#if helperProp}
<HelperDialogs helperType={helperProp} callback={() => (helperProp = undefined)}></HelperDialogs>
{/if}