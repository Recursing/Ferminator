<script lang="ts">
  import { random } from "lodash";
  import {
    type Grid,
    gridToGuesstimateScript,
    xlsxToGrid,
    gridToMermaidGraph,
  } from "./sheetToGuesstimate";

  import mermaid from "mermaid";
  import { onMount } from "svelte";
  import pako from "pako";
  import { Buffer } from "buffer";

  let files: FileList;
  let guesstimateScript = "";
  let mermaidSvg = "";
  let krokiEncoded = "";

  const generateScript = async (file: File) => {
    console.log(`Loading ${file.name}: ${file.size} bytes`);
    const grid: Grid = await xlsxToGrid(await file.arrayBuffer());
    guesstimateScript = gridToGuesstimateScript(grid);
    console.log(grid);
    const mermaidGraph = gridToMermaidGraph(grid);
    console.log(mermaidGraph);
    mermaidSvg = mermaid.mermaidAPI.render(
      "asd" + random(0, 2 ** 30),
      mermaidGraph
    );
    console.log(`Loaded ${file.name}`);

    const data = Buffer.from(mermaidGraph, "utf8");
    const compressed = pako.deflate(data, { level: 9 });
    console.log(mermaidGraph);
    krokiEncoded =
      "https://kroki.io/mermaid/svg/" +
      Buffer.from(compressed)
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_");
    console.log(krokiEncoded);
    console.log(data);
    console.log(compressed);
  };

  $: if (files && files[0]) {
    generateScript(files[0]);
  }

  var config = {
    startOnLoad: false,
    flowchart: { useMaxWidth: false, htmlLabels: true },
  };
  onMount(() => {
    mermaid.initialize(config);
  });
</script>

<span>Upload xslx:</span>
<input
  accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  type="file"
  id="myFile"
  bind:files
  name="filename"
/>

{#if guesstimateScript}
  <h2>
    Click on the text below, copy to clipboard (<kbd>ctrl+C</kbd>), open the
    guesstimate model, open browser console (<kbd>F12</kbd>), paste, enter and
    enjoy
  </h2>
  <pre style="user-select:all;">{guesstimateScript}</pre>
  <h2>
    Computation graph:
    <a href={krokiEncoded}>Link</a>
  </h2>
  <div id="graphDiv">{@html mermaidSvg}</div>
{/if}

<style>
  kbd {
    background-color: #eee;
    border-radius: 3px;
    border: 1px solid #b4b4b4;
    box-shadow: 0 1px 1px rgba(0, 0, 0, 0.2),
      0 2px 0 0 rgba(255, 255, 255, 0.7) inset;
    color: #333;
    display: inline-block;
    font-size: 0.85em;
    font-weight: 700;
    line-height: 1;
    padding: 2px 4px;
    white-space: nowrap;
  }
  h2,
  span {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
      Oxygen-Sans, Ubuntu, Cantarell, "Helvetica Neue", sans-serif;
    margin-bottom: 0;
  }
  pre {
    border-bottom: solid 1px #ccc;
    padding-bottom: 1rem;
    margin-bottom: 1rem;
    overflow-x: hidden;
  }
</style>
