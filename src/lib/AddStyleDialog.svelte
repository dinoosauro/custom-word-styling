<script lang="ts">
    import { cubicInOut } from "svelte/easing";
    import { fade } from "svelte/transition";
    import { lang } from "../Scripts/Language";

    let name = "";
    let type = "Paragraph";
    let quickStyle = true;
    const { callback }: { callback: (name: string, type: string, quickStyle: boolean) => void } =
        $props();
</script>

<div
    class="dialog"
    in:fade={{ duration: 400, easing: cubicInOut }}
    out:fade={{ duration: 400, easing: cubicInOut }}
>
    <div>
        <h2>{lang("Add new style:")}"</h2>
        <label class="flex hcenter gap">
            {lang("Style name:")} <input type="text" bind:value={name} />
        </label><br />
        <label class="flex hcenter gap">
            {lang("Style type:")} <div class="selectContainer">
                <select bind:value={type}>
                    {#each [["Paragraph", lang("Paragraph")], ["Character", lang("Character")], ["List", lang("List")], ["Table", lang("Table")]] as [char, display]}
                        <option value={char}>{display}</option>
                    {/each}
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            <input type="checkbox" bind:checked={quickStyle}>
            {lang("Show this style as a quick style in Word's \"Styles\" tab")}
        </label><br>
        <button onclick={() => callback(name, type, quickStyle)}>Save</button><br /><br />
    </div>
</div>

<style>
    input,
    select {
        width: 100%;
    }
</style>
