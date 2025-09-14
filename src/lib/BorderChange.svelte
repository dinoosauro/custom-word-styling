<script lang="ts">
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    import DeleteButton from "./DeleteButton.svelte";

    const {
        sourceBorder,
    }: {
        sourceBorder: Word.BorderCollection;
    } = $props();
    let specificBorderIndex = $state("0");
    const availablePoints = ["None","Pt025","Pt050","Pt075","Pt100","Pt150","Pt225","Pt300","Pt450","Pt600"];
    const availableStyles = ["None","Single","Double","Dotted","Dashed","DotDashed","Dot2Dashed","Triple","ThinThickSmall","ThickThinSmall","ThinThickThinSmall","ThinThickMed","ThickThinMed","ThinThickThinMed","ThinThickLarge","ThickThinLarge","ThinThickThinLarge","Wave","DoubleWave","DashedSmall","DashDotStroked","ThreeDEmboss","ThreeDEngrave"];
</script>

{#each ["Inside", "Outside"] as key}
    <Card secondCard={true}>
        <h3>{lang(`${key} border`)}:</h3>
        <label class="flex hcenter gap">
            {lang("Color")}: <input
                type="color"
                value={sourceBorder[
                    `${key === "Inside" ? "inside" : "outside"}BorderColor`
                ]}
                onchange={(e) =>
                    (sourceBorder[
                        `${key === "Inside" ? "inside" : "outside"}BorderColor`
                    ] = e.currentTarget.value)}
            />
            <DeleteButton
                callback={(e) => {
                    // @ts-ignore
                    delete sourceBorder[
                        `${key === "Inside" ? "inside" : "outside"}BorderColor`
                    ];
                    const possibleItem = (e.target as HTMLElement)
                        .closest("label")
                        ?.querySelector("input[type=color]");
                    if (possibleItem)
                        (possibleItem as HTMLInputElement).value = "#ffffff";
                }}
            ></DeleteButton>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Type")}: <div class="selectContainer">
                <select
                    bind:value={
                        sourceBorder[
                            `${key === "Inside" ? "inside" : "outside"}BorderType`
                        ]
                    }
                >
                    {#each availableStyles as option}
                        <option value={option}>{option}</option>
                    {/each}
                </select>
            </div>
        </label><br />
        <label class="flex hcenter gap">
            {lang("Width")}: <div class="selectContainer">
                <select
                    value={sourceBorder[
                        `${key === "Inside" ? "inside" : "outside"}BorderWidth`
                    ]}
                    onchange={(e) => {
                        sourceBorder[
                            `${key === "Inside" ? "inside" : "outside"}BorderWidth`
                        ] = e.currentTarget.value as "Mixed";
                    }}
                >
                    {#each availablePoints as option}
                    <option value={option}>{option.startsWith("Pt") ? `${option[2]}.${option.substring(3)} points` : option}</option>
                    {/each}
                </select>
            </div>
        </label>
    </Card>
    <br />
{/each}
{#if sourceBorder.items?.length > 0}
    <Card secondCard={true}>
        <h3>{lang("Specific border options")}:</h3>
        <div class="selectContainer">
            <select bind:value={specificBorderIndex}>
                {#each sourceBorder.items as borderItem, i}
                    <option value={i.toString()}>{lang(`${borderItem.location} border${borderItem.location === "All" ? "s" : ""}`)}</option>
                {/each}
            </select>
        </div>
        <br />
        <label class="flex hcenter gap">
            <input type="checkbox" bind:checked={sourceBorder.items[+specificBorderIndex].visible}>The border should be visible
        </label><br>
        <label class="flex hcenter gap">
            {lang("Color")}: <input
                type="color"
                value={sourceBorder.items[+specificBorderIndex].color ||
                    "#ffffff"}
                onchange={(e) =>
                    (sourceBorder.items[+specificBorderIndex].color =
                        e.currentTarget.value)}
            />
            <DeleteButton
                callback={(e) => {
                    // @ts-ignore
                    delete sourceBorder.items[+specificBorderIndex].color;
                    const possibleItem = (e.target as HTMLElement)
                        .closest("label")
                        ?.querySelector("input[type=color]");
                    if (possibleItem)
                        (possibleItem as HTMLInputElement).value = "#000000";
                }}
            ></DeleteButton>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Border width")}: <select bind:value={sourceBorder.items[+specificBorderIndex].width}>
                {#each availablePoints as option}
                <option value={option}>{option.startsWith("Pt") ? `${option[2]}.${option.substring(3)} points` : option}</option>
                {/each}
            </select>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Border type")}: <select bind:value={sourceBorder.items[+specificBorderIndex].type}>
                {#each availableStyles as option}
                    <option value={option}>{option}</option>
                {/each}
            </select>
        </label><br>
   </Card>
{/if}
