<script lang="ts">
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    const {data}: {
        data: {
            gradientDegrees: number,
            currentColor: string,
            /**
             * [Color position, hex code, transparency]
             */
            colors: [number, string, number][],
            currentPercentage: number,
            currentTransparency?: number
        }
    } = $props();
    let colorMaps = $state([]) as [number, string, number][];
    $effect(() => {
        data.colors = colorMaps;
    })
</script>
<label class="flex hcenter gap">
    {lang("Gradient angle")}: <input min="0" max="360" type="number" bind:value={data.gradientDegrees}>
</label><br>
<Card>
    <h3>{lang("Add colors")}:</h3>
    <label class="flex hcenter gap">
        {lang("Color")}: <input type="color" bind:value={data.currentColor}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Transparency")}: <input type="range" min="0" max="1" step="0.001" bind:value={data.currentTransparency}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Position of this color (between 0 and 1)")}: <input type="number" max="1" min="0" step="0.001" bind:value={data.currentPercentage}>
    </label><br>
    <button onclick={() => (colorMaps = [...colorMaps, [data.currentPercentage, data.currentColor, data.currentTransparency ?? 0]])}>Aggiungi colore</button><br><br>
    <Card secondCard={true}>
        <div class="flex gap" style="flex-wrap: wrap">
            {#each colorMaps as color, i}
            <span class="flex hcenter gap" onclick={() => (colorMaps = [...colorMaps.slice(0, i), ...colorMaps.slice(i + 1)])}>
                <div style={`width: 20px; height: 20px; background-color: ${color[1]}`}></div>
                <span>{color[1]} ({color[0] * 100}% {lang("position")} | {color[2] * 100}% {lang("transparency")})</span>
            </span>
            {/each}
        </div>
    </Card>
</Card>