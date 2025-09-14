<script lang="ts">
    import { cubicInOut } from "svelte/easing";
    import { fade } from "svelte/transition";
    import { lang } from "../Scripts/Language";
    import Card from "./Card.svelte";
    let selectedLicense = $state("Svelte");
    interface LicenseInfo {
        url: string,
        authors: string,
    }
    const authorMap = new Map<string, LicenseInfo>([
        ["Svelte", {url: "https://github.com/sveltejs/svelte", authors: "2016-2025 Svelte Contributors"}],
        ["Vite", {url: "https://github.com/vitejs/vite", authors: "2019-present, VoidZero Inc. and Vite contributors"}],
        ["Fluent UI System Icons", {url: "https://github.com/microsoft/fluentui-system-icons", authors: "2020 Microsoft Corporation"}],
        ["Custom Word Style", {url: "https://github.com/dinoosauro/custom-word-style", authors: "2025 Dinoosauro"}]
    ])
    const {callback}: {callback: () => void} = $props();
</script>
<div
    class="dialog"
    in:fade={{ duration: 400, easing: cubicInOut }}
    out:fade={{ duration: 400, easing: cubicInOut }}
>
    <div>
        <h2>{lang("Open source licenses")}:</h2>
        <p>{lang("Select one of the following open source libraries to see its license")}.</p>
        <div class="selectContainer">
            <select bind:value={selectedLicense}>
                <option value="Svelte">Svelte</option>
                <option value="Vite">Vite</option>
                <option value="Fluent UI System Icons">Fluent UI System Icons</option>
                <option value="Custom Word Style">Custom Word Style</option>
            </select>
        </div><br>
        <Card secondCard={true}>
            <h3><u onclick={() => Office.isSetSupported("OpenBrowserWindowApi", "1.1") && Office.context.ui.openBrowserWindow(authorMap.get(selectedLicense)?.url as string)}>{selectedLicense}</u></h3>    
            {#if !Office.isSetSupported("OpenBrowserWindowApi", "1.1")}
            <i>{authorMap.get(selectedLicense)?.url}</i><br><br>
            {/if}
            <p>
                MIT License<br><br>
                Copyright (c) {authorMap.get(selectedLicense)?.authors}<br><br>
                Permission is hereby granted, free of charge, to any person obtaining a copy
                of this software and associated documentation files (the "Software"), to deal
                in the Software without restriction, including without limitation the rights
                to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
                copies of the Software, and to permit persons to whom the Software is
                furnished to do so, subject to the following conditions:<br><br>

                The above copyright notice and this permission notice shall be included in all
                copies or substantial portions of the Software.<br><br>

                THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
                IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
                FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
                AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
                LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
                OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
                SOFTWARE.
            </p>
        </Card><br>
        <button onclick={() => callback()}>{lang("Close")}</button>
    </div>
</div>
