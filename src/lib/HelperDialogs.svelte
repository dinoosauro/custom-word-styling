<script lang="ts">
    import { fade } from "svelte/transition";
    import type { HelperType } from "../Scripts/HelperType";
    import { cubicInOut } from "svelte/easing";
    import CellSpacing from "../assets/CellSpacing.jpg"
    import IndentNegative from "../assets/IndentNegative.jpg"
    import IntentPositive from "../assets/IndentPositive.jpg"
    import ListNumberFormat from "../assets/ListNumberFormat.jpg"
    import ListSecondLinePosition from "../assets/ListSecondLineTextPosition.jpg"
    import ListSecondLinePosition2 from "../assets/ListSecondLineTextPosition2.jpg"
    import ParagraphSameLine from "../assets/ParagraphSameLine.gif"
    import ParagraphSamePage from "../assets/ParagraphSamePage.gif"
    import Shading from "../assets/Shading.jpg"
    import ListTabPosition from "../assets/ListTabPosition.jpg"
    import { lang } from "../Scripts/Language";
    const {helperType, callback}: {helperType: HelperType, callback: () => void} = $props();
</script>

<div
    class="dialog"
    in:fade={{ duration: 400, easing: cubicInOut }}
    out:fade={{ duration: 400, easing: cubicInOut }}
>
    <div>
        {#if helperType === "CellSpacing"}
            <h2>{lang("Table cell spacing")}:</h2>
            <p>{lang("You can choose how much empty space there should be around each cell.")}<br><b>{lang("Value unit")}:</b> {lang("points")}</p><br>
            <img alt={lang("Table with some empty space around each cell")} src={CellSpacing}>
        {:else if helperType === "Indent"}
            <h2>{lang("Paragraph indent")}:</h2>
            <p>{lang("By changing the paragraph indent value, you can put some empty space before the first line. You can both put a positive value (image 1) or a negative value (image 2)")}.</p><br><b>{lang("Value unit")}:</b> {lang("points")}<br>
            <img alt={lang("Paragraph with indent positive applied: the text of the first line starts after if compared to the text of the other lines")} src={IntentPositive}><br>
            <img alt={lang("Paragraph with indent negative applied: the text of the first line starts before the text of the other lines")} src={IndentNegative}>
        {:else if helperType === "ListNumberFormat"}
            <h2>{lang("List number format")}:</h2>
            <p>{lang("You can change how the numbers or bullet points should be displayed, by adding some text before or after them. In the image example, the")} <code> - Example)</code> {lang("text has been added after the number")}.<br><b>{lang("Value unit")}:</b> {lang("points")}</p>
            <img alt={lang("Custom number format")} src={ListNumberFormat}>
        {:else if helperType === "ListSecondLine"}
            <h2>{lang("Second line position")}:</h2>
            <p>{lang("With this property, you can change the space before the start of every line after the first one. For example, in the first image this property was set to a normal value, while in the second one was set to")} <code>40</code>.<br><b>{lang("Value unit")}:</b> {lang("points")}</p>
            <img alt={lang("Normal second line position property")} src={ListSecondLinePosition}>
            <img alt={lang("Second line position set to 40: the text of the second and third line is positioned more at the right than the first line")} src={ListSecondLinePosition2}>   
        {:else if helperType === "ParagraphSameLine"}
            <h2>{lang("Keep paragraph on the same line")}:</h2>
            <p>{lang("As the title says, if this option is enabled, the paragraph will be put in the same page as the paragraph which follows it. Note that, in the video, the last two lines of the Lorem ipsum have been made a new line")}.</p><br>
            <img alt={lang("GIF where the entire paragraph goes to a new page")} src={ParagraphSameLine}>
        {:else if helperType === "ParagraphSamePage"}
            <h2>{lang("Keep paragraph on the same page")}:</h2>
            <p>{lang("If enabled, all the lines of the paragraph will be kept in the same page. So, if at least one line should go in a new page, all the paragraph will be put in the new page")}.</p><br>
            <img alt={lang("GIF where the second last paragraph goes to the new page with the last paragraph")} src={ParagraphSamePage}>
        {:else if helperType === "Shading"}
            <h2>{lang("Shading options")}:</h2>
            <p>{lang("With shading options, you can customize the background color of the paragraph. Here you can see an example, done by mixing red and yellow and by setting the shading type to")} <code>40%</code>.</p><br>
            <img alt={lang("Custom shading")} src={Shading}>
        {:else if helperType === "LevelOutline"}
            <h2>{lang("Level outline")}:</h2>
            <p>{lang("The level the paragraph should be displayed if a summary is generated, from 1 to 9. If it shouldn't be displayed,")} <code>BodyText</code>{lang(" should be selected")}.</p>
        {:else if helperType === "ListPosition"}
            <h2>{lang("Tab position")}:</h2>
            <p>{lang("Choose how much tabulation space should be added between the number/bullet point and the first character of the list. If you haven't set it, you might see a really high value. This is normal, but know that usually setting it of a few tens of points is more than enough (since setting it to a high value, like thousands of points, will cause Word to reject the style change")}.<br><b>{lang("Value unit")}:</b> {lang("points")}</p><br>
            <img alt={lang("First item on a list, with a lot of spcae between the number and the first character")} src={ListTabPosition}>
        {/if}<br>
        <button onclick={() => callback()}>{lang("Close")}</button>
</div>
</div>

<style>
    img, video {
        width: 100%;
        height: auto;
        border: 1px solid var(--text);
        display: block;
        border-radius: 8px;
    }
</style>