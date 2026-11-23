"use client";

import { useEffect, useRef } from "react";
import type { StudentRecord } from "@/lib/matcher";

type Rel = {
  id: string;
  left: StudentRecord;
  right: StudentRecord;
  score: number;
  reasons: string[];
  autoConfirmEligible: boolean;
  status: string;
};

function label(record: Rel["left"]) {
  return String(record.data.Name ?? record.data["Full Name"] ?? record.data.name ?? record.data["Roll No"] ?? "Unnamed record");
}
function hash(value: string) {
  let h = 2166136261;
  for (let i = 0; i < value.length; i++) h = Math.imul(h ^ value.charCodeAt(i), 16777619);
  return h >>> 0;
}

export default function RelationshipScene({ relationships, mode, onInspect }: {
  relationships: Rel[];
  mode: "all" | "confirmed" | "pending" | "rejected";
  onInspect: (relationship: Rel | null) => void;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const inspectRef = useRef(onInspect);
  inspectRef.current = onInspect;

  useEffect(() => {
    let disposed = false;
    let cleanup = () => {};
    const host = hostRef.current;
    if (!host) return;

    const visible = relationships.filter((r) => mode === "all" || (mode === "confirmed"
      ? r.status === "Confirmed same person" || r.status === "Auto-confirmed same person"
      : mode === "rejected" ? r.status === "Not a match"
      : r.status !== "Confirmed same person" && r.status !== "Auto-confirmed same person" && r.status !== "Not a match"));
    const plotted = visible.slice(0, 3500);

    void import("three").then((THREE) => {
      if (disposed || !host) return;
      const scene = new THREE.Scene();
      scene.background = new THREE.Color("#07110d");
      const camera = new THREE.PerspectiveCamera(48, host.clientWidth / Math.max(1, host.clientHeight), 0.1, 500);
      camera.position.set(0, 0.2, 19);
      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: false, powerPreference: "high-performance" });
      renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.8));
      renderer.setSize(host.clientWidth, host.clientHeight);
      renderer.outputColorSpace = THREE.SRGBColorSpace;
      host.appendChild(renderer.domElement);

      scene.add(new THREE.AmbientLight(0xffffff, 1.2));
      const keyLight = new THREE.PointLight(0xb6ffd4, 24, 45);
      keyLight.position.set(0, 5, 9);
      scene.add(keyLight);

      const d1Map = new Map<string, { record: Rel["left"]; position: InstanceType<typeof THREE.Vector3>; links: number[] }>();
      const d2Map = new Map<string, { record: Rel["right"]; position: InstanceType<typeof THREE.Vector3>; links: number[] }>();
      const d1Ids: string[] = [];
      const d2Ids: string[] = [];
      plotted.forEach((r, i) => {
        if (!d1Map.has(r.left.id)) { d1Map.set(r.left.id, { record: r.left, position: new THREE.Vector3(), links: [] }); d1Ids.push(r.left.id); }
        if (!d2Map.has(r.right.id)) { d2Map.set(r.right.id, { record: r.right, position: new THREE.Vector3(), links: [] }); d2Ids.push(r.right.id); }
        d1Map.get(r.left.id)!.links.push(i);
        d2Map.get(r.right.id)!.links.push(i);
      });
      const arrange = (ids: string[], map: typeof d1Map, side: number) => ids.forEach((id, i) => {
        const item = map.get(id)!;
        const h = hash(id);
        const angle = (i / Math.max(1, ids.length)) * Math.PI * 2 + (h % 1000) / 1000 * 0.12;
        const radius = 2.2 + ((h >>> 8) % 1000) / 1000 * 2.4;
        const y = Math.sin(angle) * radius;
        const z = Math.cos(angle) * radius * 0.62 + (((h >>> 18) % 100) / 100 - 0.5) * 1.8;
        item.position.set(side * (5.2 + Math.cos(angle) * 0.65), y, z);
      });
      arrange(d1Ids, d1Map, -1);
      arrange(d2Ids, d2Map, 1);

      const edges: number[] = [];
      const edgeColors: number[] = [];
      const colorFor = (status: string) => status === "Confirmed same person" || status === "Auto-confirmed same person" ? new THREE.Color("#baffd5") : status === "Not a match" ? new THREE.Color("#ff7185") : new THREE.Color("#8dd6ae");
      plotted.forEach((r, i) => {
        const a = d1Map.get(r.left.id)!.position;
        const b = d2Map.get(r.right.id)!.position;
        edges.push(a.x, a.y, a.z, b.x, b.y, b.z);
        const c = colorFor(r.status);
      });
      const edgeGeometry = new THREE.BufferGeometry();
      edgeGeometry.setAttribute("position", new THREE.Float32BufferAttribute(edges, 3));
      const initialEdgeColors = plotted.flatMap((r) => { const c = colorFor(r.status); return [c.r * 0.78, c.g * 0.78, c.b * 0.78, c.r * 0.78, c.g * 0.78, c.b * 0.78]; });
      edgeGeometry.setAttribute("color", new THREE.Float32BufferAttribute(initialEdgeColors, 3));
      // Use Three.js' built-in line material for dependable visibility across GPUs.
      // A separate moving-points layer supplies the heartbeat travelling along each edge.
      const edgeMaterial = new THREE.LineBasicMaterial({
        vertexColors: true,
        transparent: true,
        opacity: 0.58,
        depthWrite: false,
        blending: THREE.AdditiveBlending,
      });
      const lineSegments = new THREE.LineSegments(edgeGeometry, edgeMaterial);
      lineSegments.userData.kind = "edges";
      scene.add(lineSegments);

      const pulseCount = plotted.length * 2;
      const pulsePositions = new Float32Array(pulseCount * 3);
      const pulseColors = new Float32Array(pulseCount * 3);
      for (let i = 0; i < plotted.length; i++) {
        const c = colorFor(plotted[i].status);
        for (let beat = 0; beat < 2; beat++) {
          const k = (i * 2 + beat) * 3;
          pulseColors[k] = 0.72;
          pulseColors[k + 1] = 1.0;
          pulseColors[k + 2] = 0.84;
        }
      }
      const pulseGeometry = new THREE.BufferGeometry();
      pulseGeometry.setAttribute("position", new THREE.BufferAttribute(pulsePositions, 3));
      pulseGeometry.setAttribute("color", new THREE.BufferAttribute(pulseColors, 3));
      const pulseMaterial = new THREE.PointsMaterial({
        size: 0.18,
        vertexColors: true,
        transparent: true,
        opacity: 0.98,
        depthWrite: false,
        sizeAttenuation: true,
        blending: THREE.AdditiveBlending,
      });
      const pulsePoints = new THREE.Points(pulseGeometry, pulseMaterial);
      pulsePoints.userData.kind = "pulses";
      scene.add(pulsePoints);

      const nodeMeshes: Array<{ mesh: InstanceType<typeof THREE.Mesh>; ids: string[]; side: "d1" | "d2" }> = [];
      const makeNodes = (ids: string[], map: typeof d1Map, side: "d1" | "d2") => {
        const geometry = new THREE.SphereGeometry(0.075, 10, 10);
        const material = new THREE.MeshStandardMaterial({ color: side === "d1" ? "#a8f2c4" : "#dcffe9", emissive: side === "d1" ? "#245c3d" : "#37684d", emissiveIntensity: 0.65, roughness: 0.35, metalness: 0.1 });
        const mesh = new THREE.InstancedMesh(geometry, material, ids.length);
        mesh.userData.kind = side;
        const dummy = new THREE.Object3D();
        ids.forEach((id, i) => { dummy.position.copy(map.get(id)!.position); const scale = 0.75 + Math.min(1.8, map.get(id)!.links.length * 0.16); dummy.scale.setScalar(scale); dummy.updateMatrix(); mesh.setMatrixAt(i, dummy.matrix); });
        mesh.instanceMatrix.needsUpdate = true;
        scene.add(mesh);
        nodeMeshes.push({ mesh, ids, side });
      };
      makeNodes(d1Ids, d1Map, "d1");
      makeNodes(d2Ids, d2Map, "d2");

      // Fine radial guide arcs and subtle particle field create the observatory aesthetic.
      const ring = new THREE.Mesh(new THREE.TorusGeometry(3.25, 0.006, 4, 180), new THREE.MeshBasicMaterial({ color: "#5c8d70", transparent: true, opacity: 0.22 }));
      ring.rotation.y = Math.PI / 2;
      scene.add(ring);
      const particleCount = 850;
      const particlePositions = new Float32Array(particleCount * 3);
      for (let i = 0; i < particleCount; i++) {
        const h = hash(`dust-${i}`); const angle = (h % 10000) / 10000 * Math.PI * 2;
        const radius = 2.5 + ((h >>> 10) % 1000) / 1000 * 6.5;
        particlePositions[i * 3] = Math.cos(angle) * radius;
        particlePositions[i * 3 + 1] = Math.sin(angle) * radius * 0.55;
        particlePositions[i * 3 + 2] = (((h >>> 20) % 1000) / 1000 - 0.5) * 7;
      }
      const dustGeometry = new THREE.BufferGeometry(); dustGeometry.setAttribute("position", new THREE.BufferAttribute(particlePositions, 3));
      scene.add(new THREE.Points(dustGeometry, new THREE.PointsMaterial({ color: "#b8e8c9", size: 0.018, transparent: true, opacity: 0.42, sizeAttenuation: true })));

      const raycaster = new THREE.Raycaster(); raycaster.params.Line = { threshold: 0.12 };
      const pointer = new THREE.Vector2();
      let activeNode: { side: "d1" | "d2"; id: string } | null = null;
      let downX = 0; let downY = 0; let dragging = false;
      let rotationY = 0; let rotationX = 0; let zoom = 19;
      const setHighlight = (index: number | null, node: typeof activeNode) => {
        activeNode = node;
        const colors: number[] = [];
        plotted.forEach((r, i) => {
          const related = index === i || (node && (node.side === "d1" ? r.left.id === node.id : r.right.id === node.id));
          const c = colorFor(r.status); const alpha = related ? 1 : (index !== null || node ? 0.035 : 0.45);
          colors.push(c.r * alpha, c.g * alpha, c.b * alpha, c.r * alpha, c.g * alpha, c.b * alpha);
        });
        edgeGeometry.setAttribute("color", new THREE.Float32BufferAttribute(colors, 3));
        edgeGeometry.attributes.color.needsUpdate = true;
        const pulseTint: number[] = [];
        plotted.forEach((r, i) => {
          const related = index === i || (node && (node.side === "d1" ? r.left.id === node.id : r.right.id === node.id));
          const alpha = related ? 1 : (index !== null || node ? 0.025 : 1);
          for (let beat = 0; beat < 2; beat++) pulseTint.push(0.72 * alpha, 1.0 * alpha, 0.84 * alpha);
        });
        pulseGeometry.setAttribute("color", new THREE.Float32BufferAttribute(pulseTint, 3));
        pulseGeometry.attributes.color.needsUpdate = true;
        pulseMaterial.opacity = (index !== null || node) ? 0.95 : 0.88;
        edgeMaterial.opacity = (index !== null || node) ? 0.38 : 0.58;
        nodeMeshes.forEach(({ mesh, ids, side }) => ids.forEach((id, i) => {
          const isHot = node?.side === side && node.id === id;
          const isConnected = index !== null && (side === "d1" ? plotted[index]?.left.id === id : plotted[index]?.right.id === id);
          const matrix = new THREE.Matrix4(); mesh.getMatrixAt(i, matrix);
          const scale = isHot || isConnected ? 1.9 : (node || index !== null ? 0.55 : 1);
          const pos = new THREE.Vector3().setFromMatrixPosition(matrix);
          const obj = new THREE.Object3D(); obj.position.copy(pos); obj.scale.setScalar(scale); obj.updateMatrix(); mesh.setMatrixAt(i, obj.matrix);
        }));
        nodeMeshes.forEach(({ mesh }) => { mesh.instanceMatrix.needsUpdate = true; });
      };
      const onPointerMove = (event: PointerEvent) => {
        const rect = renderer.domElement.getBoundingClientRect();
        pointer.x = ((event.clientX - rect.left) / rect.width) * 2 - 1;
        pointer.y = -((event.clientY - rect.top) / rect.height) * 2 + 1;
        if (dragging) { rotationY += (event.clientX - downX) * 0.004; rotationX += (event.clientY - downY) * 0.003; downX = event.clientX; downY = event.clientY; return; }
        raycaster.setFromCamera(pointer, camera);
        const hits = raycaster.intersectObjects([...nodeMeshes.map(n => n.mesh), lineSegments], false);
        const hit = hits[0];
        if (!hit) { if (activeNode) { setHighlight(null, null); inspectRef.current(null); } return; }
        if (hit.object === lineSegments && hit.index !== undefined) {
          const index = Math.floor(hit.index / 2);
          if (index >= 0 && index < plotted.length) { setHighlight(index, null); inspectRef.current(plotted[index]); }
        } else if (hit.object instanceof THREE.InstancedMesh && hit.instanceId !== undefined) {
          const entry = nodeMeshes.find(n => n.mesh === hit.object); if (!entry) return;
          const id = entry.ids[hit.instanceId]; const node = { side: entry.side, id } as const;
          const related = (entry.side === "d1" ? d1Map : d2Map).get(id)?.links[0];
          setHighlight(related ?? null, node);
          const rel = related === undefined ? null : plotted[related]; inspectRef.current(rel);
        }
      };
      const onPointerDown = (event: PointerEvent) => { downX = event.clientX; downY = event.clientY; dragging = true; renderer.domElement.setPointerCapture(event.pointerId); };
      const onPointerUp = () => { dragging = false; };
      const onClick = (event: MouseEvent) => {
        const rect = renderer.domElement.getBoundingClientRect(); pointer.x = ((event.clientX - rect.left) / rect.width) * 2 - 1; pointer.y = -((event.clientY - rect.top) / rect.height) * 2 + 1;
        raycaster.setFromCamera(pointer, camera); const hit = raycaster.intersectObject(lineSegments, false)[0];
        if (hit?.index !== undefined) { const idx = Math.floor(hit.index / 2); if (idx >= 0 && idx < plotted.length) { setHighlight(idx, null); inspectRef.current(plotted[idx]); } }
      };
      const onWheel = (event: WheelEvent) => { zoom = Math.max(10, Math.min(30, zoom + event.deltaY * 0.008)); };
      renderer.domElement.addEventListener("pointermove", onPointerMove);
      renderer.domElement.addEventListener("pointerdown", onPointerDown);
      renderer.domElement.addEventListener("pointerup", onPointerUp);
      renderer.domElement.addEventListener("pointercancel", onPointerUp);
      renderer.domElement.addEventListener("click", onClick);
      renderer.domElement.addEventListener("wheel", onWheel, { passive: true });
      const onResize = () => { if (!host) return; camera.aspect = host.clientWidth / Math.max(1, host.clientHeight); camera.updateProjectionMatrix(); renderer.setSize(host.clientWidth, host.clientHeight); };
      window.addEventListener("resize", onResize);
      let frame = 0;
      const clock = new THREE.Clock();
      const animate = () => {
        frame = requestAnimationFrame(animate);
        scene.rotation.y += (rotationY - scene.rotation.y) * 0.06;
        scene.rotation.x += (rotationX - scene.rotation.x) * 0.06;
        camera.position.z += (zoom - camera.position.z) * 0.08;
        const elapsed = clock.getElapsedTime();
        const pulseAttr = pulseGeometry.getAttribute("position");
        for (let i = 0; i < plotted.length; i++) {
          const r = plotted[i];
          const a = d1Map.get(r.left.id)!.position;
          const b = d2Map.get(r.right.id)!.position;
          // Two close pulses travel D1 → D2, followed by a short resting beat.
          // Stagger each edge slightly so the whole network feels alive, not synchronized.
          const cycle = 2.5;
          const travel = 0.95;
          const phase = (elapsed + (i * 0.61803398875) % cycle) % cycle;
          const firstActive = phase < travel;
          const secondPhase = phase - 0.14;
          const secondActive = secondPhase >= 0 && secondPhase < travel;
          if (firstActive) {
            const t = phase / travel;
            pulseAttr.setXYZ(i * 2,
              a.x + (b.x - a.x) * t,
              a.y + (b.y - a.y) * t,
              a.z + (b.z - a.z) * t);
          } else pulseAttr.setXYZ(i * 2, 0, -1000, 0);
          if (secondActive) {
            const t = secondPhase / travel;
            pulseAttr.setXYZ(i * 2 + 1,
              a.x + (b.x - a.x) * t,
              a.y + (b.y - a.y) * t,
              a.z + (b.z - a.z) * t);
          } else pulseAttr.setXYZ(i * 2 + 1, 0, -1000, 0);
        }
        pulseAttr.needsUpdate = true;
        renderer.render(scene, camera);
      };
      animate();
      cleanup = () => { cancelAnimationFrame(frame); window.removeEventListener("resize", onResize); renderer.domElement.removeEventListener("pointermove", onPointerMove); renderer.domElement.removeEventListener("pointerdown", onPointerDown); renderer.domElement.removeEventListener("pointerup", onPointerUp); renderer.domElement.removeEventListener("pointercancel", onPointerUp); renderer.domElement.removeEventListener("click", onClick); renderer.domElement.removeEventListener("wheel", onWheel); scene.traverse(obj => { if (obj instanceof THREE.Mesh || obj instanceof THREE.Points || obj instanceof THREE.LineSegments) { obj.geometry.dispose(); const mat = obj.material; if (Array.isArray(mat)) mat.forEach(m => m.dispose()); else mat.dispose(); } }); renderer.dispose(); renderer.domElement.remove(); };
    }).catch((error) => { console.error("Three.js visualization failed to initialize", error); });
    return () => { disposed = true; cleanup(); };
  }, [relationships, mode]);

  return <div ref={hostRef} className="absolute inset-0 touch-none" aria-label="Interactive Three.js relationship network"><div className="pointer-events-none absolute bottom-5 left-1/2 z-10 -translate-x-1/2 whitespace-nowrap rounded-full border border-emerald-200/10 bg-black/35 px-4 py-2 text-[9px] tracking-[.16em] text-emerald-100/40 backdrop-blur">DRAG TO ROTATE · SCROLL TO ZOOM · HOVER TO TRACE CONNECTIONS</div></div>;
}
