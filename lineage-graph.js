/**
 * Lineage Graph Renderer — Canvas 2D
 *
 * Draws the data dependency graph produced by ASTEngine.analyzeLineage().
 * Uses a layered (hierarchical) layout: columns → variables → transforms → comparisons → outputs.
 *
 * Color palette matches the gold-on-black theme of the application.
 */
const LineageGraph = (() => {
  'use strict';

  const COLORS = {
    column:     { bg: '#d4af37', text: '#0a0b0d', border: '#f3e5ab', glow: 'rgba(212,175,55,0.35)' },
    variable:   { bg: '#1a1c23', text: '#f0f6fc', border: '#484f58', glow: 'rgba(139,148,158,0.15)' },
    transform:  { bg: '#15202b', text: '#79c0ff', border: '#1f6feb', glow: 'rgba(31,111,235,0.2)' },
    comparison: { bg: '#2d1a12', text: '#ffa657', border: '#d4711a', glow: 'rgba(212,113,26,0.2)' },
    output:     { bg: '#0d2117', text: '#3fb950', border: '#238636', glow: 'rgba(35,134,54,0.25)' }
  };

  const EDGE_COLORS = {
    acquisition: '#d4af37',
    access:      '#8b949e',
    transform:   '#79c0ff',
    comparison:  '#ffa657',
    control:     '#f85149',
    output:      '#3fb950'
  };

  const NODE_W = 150;
  const NODE_H = 38;
  const LAYER_GAP_X = 220;
  const NODE_GAP_Y = 56;
  const PAD = 50;

  /**
   * Renders the lineage graph onto a canvas element.
   * @param {HTMLCanvasElement} canvas
   * @param {{ nodes: Array, edges: Array }} graphData
   */
  function render(canvas, graphData) {
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const layers = {
      column: [],
      variable: [],
      transform: [],
      comparison: [],
      output: []
    };

    // Assign nodes to layers
    graphData.nodes.forEach(node => {
      if (layers[node.type]) {
        layers[node.type].push(node);
      }
    });

    // Layout planning: Calculate X, Y for each node
    const positions = {};
    let currentX = PAD;

    const layerOrder = ['column', 'variable', 'transform', 'comparison', 'output'];
    let maxLayerHeight = 0;

    layerOrder.forEach(layerKey => {
      const nodes = layers[layerKey];
      if (nodes.length > 0) {
        const height = nodes.length * NODE_GAP_Y;
        if (height > maxLayerHeight) maxLayerHeight = height;
      }
    });

    const canvasH = Math.max(500, maxLayerHeight + PAD * 2);
    const canvasW = PAD * 2 + (layerOrder.filter(k => layers[k].length > 0).length - 1) * LAYER_GAP_X;

    // Set canvas dimensions
    canvas.width = canvasW;
    canvas.height = canvasH;

    // Draw dark background grid pattern
    ctx.fillStyle = '#0a0b0d';
    ctx.fillRect(0, 0, canvasW, canvasH);
    drawBackgroundGrid(ctx, canvasW, canvasH);

    layerOrder.forEach(layerKey => {
      const nodes = layers[layerKey];
      if (nodes.length === 0) return;

      const layerH = nodes.length * NODE_H + (nodes.length - 1) * (NODE_GAP_Y - NODE_H);
      let startY = (canvasH - layerH) / 2;

      nodes.forEach((node, index) => {
        const y = startY + index * NODE_GAP_Y;
        positions[node.id] = { x: currentX, y };
        node.x = currentX;
        node.y = y;
      });

      currentX += LAYER_GAP_X;
    });

    // Draw connections (bezier edges)
    graphData.edges.forEach(edge => {
      const fromPos = positions[edge.from];
      const toPos = positions[edge.to];
      if (fromPos && toPos) {
        const color = EDGE_COLORS[edge.type] || '#8b949e';
        drawEdge(ctx, fromPos.x + NODE_W, fromPos.y + NODE_H / 2, toPos.x, toPos.y + NODE_H / 2, color, edge);
      }
    });

    // Draw nodes
    graphData.nodes.forEach(node => {
      const pos = positions[node.id];
      if (pos) {
        drawNode(ctx, pos.x, pos.y, node);
      }
    });
  }

  function drawBackgroundGrid(ctx, w, h) {
    ctx.strokeStyle = 'rgba(212, 175, 55, 0.02)';
    ctx.lineWidth = 1;
    const step = 20;
    for (let x = 0; x < w; x += step) {
      ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, h); ctx.stroke();
    }
    for (let y = 0; y < h; y += step) {
      ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(w, y); ctx.stroke();
    }
  }

  function drawNode(ctx, x, y, node) {
    const style = COLORS[node.type] || { bg: '#21262d', text: '#c9d1d9', border: '#30363d', glow: 'transparent' };

    // Outer Glow / Shadow
    ctx.shadowColor = style.glow;
    ctx.shadowBlur = 10;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;

    // Background card
    ctx.fillStyle = style.bg;
    roundedRect(ctx, x, y, NODE_W, NODE_H, 6);
    ctx.fill();

    // Reset shadow
    ctx.shadowBlur = 0;

    // Border
    ctx.strokeStyle = style.border;
    ctx.lineWidth = 1.2;
    ctx.stroke();

    // Node details
    ctx.fillStyle = style.text;
    ctx.font = '500 11px Inter, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';

    const padding = 10;
    let label = node.label;
    const maxLen = 18;
    if (label.length > maxLen) label = label.substring(0, maxLen - 1) + '…';

    ctx.fillText(label, x + padding, y + NODE_H / 2);

    // Line number indicator (if present)
    if (node.line) {
      ctx.fillStyle = 'rgba(212, 175, 55, 0.6)';
      ctx.font = '9px SF Mono, monospace';
      ctx.textAlign = 'right';
      ctx.fillText('L' + node.line, x + NODE_W - padding, y + NODE_H / 2);
    }
  }

  function drawEdge(ctx, x1, y1, x2, y2, color, edge) {
    ctx.strokeStyle = color;
    ctx.lineWidth = edge.type === 'control' ? 1.0 : 1.6;
    if (edge.type === 'control') {
      ctx.setLineDash([4, 4]); // dashed lines for control conditions
    } else {
      ctx.setLineDash([]);
    }

    ctx.beginPath();
    ctx.moveTo(x1, y1);

    // Cubic Bezier curve for smooth flow layout routing
    const cp1x = x1 + (x2 - x1) * 0.45;
    const cp1y = y1;
    const cp2x = x1 + (x2 - x1) * 0.55;
    const cp2y = y2;

    ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, x2, y2);
    ctx.stroke();
    ctx.setLineDash([]);

    // Draw Arrowhead at destination
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    if (edge.type === 'control') {
      // Draw diamond for control gateway
      drawDiamond(ctx, x2 - 8, y2 - 4, 8, 8);
    } else {
      // standard arrow
      ctx.lineTo(x2 - 7, y2 - 4.5);
      ctx.lineTo(x2 - 5, y2);
      ctx.lineTo(x2 - 7, y2 + 4.5);
      ctx.closePath();
    }
    ctx.fill();

    // Edge label
    if (edge.label) {
      const lx = (x1 + x2) / 2;
      const ly = (y1 + y2) / 2 - 10;
      const maxLen = 16;
      const lbl = edge.label.length > maxLen ? edge.label.substring(0, maxLen - 1) + '…' : edge.label;

      // Background pill
      ctx.font = '9px Inter, sans-serif';
      const tw = ctx.measureText(lbl).width + 8;
      ctx.fillStyle = '#0a0b0d';
      ctx.globalAlpha = 0.85;
      roundedRect(ctx, lx - tw / 2, ly - 7, tw, 14, 4);
      ctx.fill();
      ctx.globalAlpha = 1.0;

      ctx.fillStyle = color;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(lbl, lx, ly);
      ctx.textBaseline = 'alphabetic';
    }
  }

  // ─── PRIMITIVES ─────────────────────────────────────────────────────
  function roundedRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
  }

  function drawDiamond(ctx, x, y, w, h) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const hw = w / 2;
    const hh = h / 2;
    ctx.beginPath();
    ctx.moveTo(cx, y - 2);
    ctx.lineTo(x + w + 4, cy);
    ctx.lineTo(cx, y + h + 2);
    ctx.lineTo(x - 4, cy);
    ctx.closePath();
  }

  return { render };
})();
