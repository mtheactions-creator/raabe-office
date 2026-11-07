// Folienzoom (D-072): PowerPoint's form (recorded by probe_powerpoint_zoom.py) is read from its Choice.
import { describe, expect, it } from "vitest";
import { alternateShapes } from "@raabe/ooxml";
import { syncZoom } from "../src/zoom";

const XML = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<mc:Choice xmlns:pslz="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom" Requires="pslz"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Folienzoom 4"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"><pslz:sldZm><pslz:sldZmObj sldId="258" cId="1"><pslz:zmPr id="{X}" returnToParent="0" transitionDur="1000"><p166:blipFill xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:blip r:embed="rId2"/></p166:blipFill><p166:spPr xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:xfrm><a:off x="0" y="0"/><a:ext cx="300" cy="400"/></a:xfrm></p166:spPr></pslz:zmPr></pslz:sldZmObj></pslz:sldZm></a:graphicData></a:graphic></p:graphicFrame></mc:Choice>
<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="5" name="Folienzoom 4"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr></p:pic></mc:Fallback></mc:AlternateContent>`;

describe("Folienzoom", () => {
  it("is read from PowerPoint's Choice, not the fallback picture", () => {
    const ac = new DOMParser().parseFromString(XML, "application/xml").documentElement;
    expect(alternateShapes(ac).map((e) => e.localName)).toEqual(["graphicFrame"]);
  });
  it("keeps the inner frame and the fallback picture with the zoom's place and size", () => {
    const ac = new DOMParser().parseFromString(XML, "application/xml").documentElement;
    const frame = alternateShapes(ac)[0];
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const xf = frame.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "xfrm")[0];
    xf.getElementsByTagNameNS(A, "off")[0].setAttribute("x", "999");
    xf.getElementsByTagNameNS(A, "ext")[0].setAttribute("cx", "600");
    syncZoom(frame);
    const exts = Array.from(ac.getElementsByTagNameNS(A, "ext")).map((e) => e.getAttribute("cx"));
    expect(exts).toEqual(["600", "600", "600"]);
    const fbOff = ac.getElementsByTagNameNS("http://schemas.openxmlformats.org/markup-compatibility/2006", "Fallback")[0].getElementsByTagNameNS(A, "off")[0];
    expect(fbOff.getAttribute("x")).toBe("999");
  });
});
