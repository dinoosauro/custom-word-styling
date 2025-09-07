<script lang="ts">
    import { lang } from "../Scripts/Language";

    const {generalStyle}: {
        generalStyle: Word.Style
    } = $props();
</script>

<label class="flex hcenter gap">
    {lang("Next paragraph style")}: <input type="text" bind:value={generalStyle.nextParagraphStyle}>
</label><br>
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={generalStyle.quickStyle}>{lang("Make this style a quick style")}
</label><br>
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={generalStyle.visibility}>{lang("Make this style visible in the \"Recommended style\" tab")}
</label><br>
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={generalStyle.unhideWhenUsed}>{lang("Make this style visible in the \"Recommended style\" tab after it has been used")}
</label><br>
<label class="flex hcenter gap">
    {lang("Priority")}: <input type="number" bind:value={generalStyle.priority}>
</label><br>
<label class="flex hcenter gap">
    {lang("Base style")}: <input type="text" bind:value={generalStyle.baseStyle}>
</label>