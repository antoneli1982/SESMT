/* SESÉ — exportação Excel de materiais em ordem alfabética + tamanho — 2026-08-12 */
(() => {
  "use strict";

  const sortedMaterialsForExcel = () => [...state.materials].sort(compareMaterialsAlphanumeric);

  function exportStockReportExcelPatched() {
    const button = document.getElementById("exportMaterialsPdf");
    if (!button) return;
    const originalLabel = button.textContent;
    button.disabled = true;
    button.textContent = "Gerando Excel...";

    try {
      if (typeof XLSX === "undefined") throw new Error("Biblioteca XLSX indisponível.");

      const materials = sortedMaterialsForExcel();
      const generatedAt = new Date().toLocaleString("pt-BR");
      const totalUnits = materials.reduce((sum, material) => sum + (Number(material.stock) || 0), 0);
      const zeroed = materials.filter(material => Number(material.stock) === 0).length;
      const critical = materials.filter(material => materialStatus(material).key === "critical").length;

      const rows = [
        ["RELATÓRIO DE ESTOQUE DE MATERIAIS", "", "", "", "", "", "", "", "", ""],
        ["Unidade", state.settings.unit || "Gestão de EPI", "", "", "", "Emissão", generatedAt, "", "", ""],
        ["Endereço", state.settings.address || "Não informado", "", "", "", "Responsável", state.settings.responsible || "Não informado", "", "", ""],
        [],
        ["ITENS CADASTRADOS", materials.length, "TOTAL EM ESTOQUE", totalUnits, "ESTOQUE CRÍTICO", critical, "SALDO ZERADO", zeroed, "", ""],
        [],
        ["#", "Código", "Descrição", "Categoria", "Tamanho", "CA", "Lote", "Localização", "Saldo", "Estoque mínimo"]
      ];

      materials.forEach((material, index) => rows.push([
        index + 1,
        material.code || "",
        upperText(material.description || ""),
        material.category || "",
        material.size || "",
        material.ca || "",
        material.lot || "",
        material.location || "",
        Number(material.stock) || 0,
        hasConfiguredMinimum(material) ? Number(material.minimum) : ""
      ]));

      rows.push(["", "", "", "", "", "", "", "TOTAL GERAL EM ESTOQUE", totalUnits, ""]);

      const sheet = XLSX.utils.aoa_to_sheet(rows);
      const headerRowIndex = 6;
      const lastRowIndex = rows.length - 1;
      sheet["!merges"] = [
        XLSX.utils.decode_range("A1:J1"),
        XLSX.utils.decode_range("B2:E2"),
        XLSX.utils.decode_range("G2:J2"),
        XLSX.utils.decode_range("B3:E3"),
        XLSX.utils.decode_range("G3:J3")
      ];
      sheet["!autofilter"] = {
        ref: XLSX.utils.encode_range(
          { r: headerRowIndex, c: 0 },
          { r: Math.max(headerRowIndex, lastRowIndex - 1), c: 9 }
        )
      };
      sheet["!freeze"] = { xSplit: 0, ySplit: headerRowIndex + 1 };
      sheet["!cols"] = [
        { wch: 7 }, { wch: 16 }, { wch: 42 }, { wch: 30 }, { wch: 14 },
        { wch: 16 }, { wch: 18 }, { wch: 28 }, { wch: 12 }, { wch: 18 }
      ];

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, "RELATÓRIO ESTOQUE");
      XLSX.writeFile(workbook, `relatorio_estoque_materiais_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast("Relatório Excel gerado em ordem alfabética e por tamanho.");
    } catch (error) {
      console.error("Falha na geração do relatório Excel", error);
      showToast("Não foi possível gerar o relatório Excel.", true);
    } finally {
      button.disabled = false;
      button.textContent = "Relatório Excel";
    }
  }

  function exportMaterialsSortedPatched() {
    try {
      if (typeof XLSX === "undefined") throw new Error("Biblioteca XLSX indisponível.");

      const rows = [[
        "Código", "Descrição", "Tamanho", "Fabricante", "Categoria", "Número do CA",
        "Validade do CA", "Lote", "Validade do lote", "Saldo", "Estoque mínimo",
        "Vida útil em dias", "Localização", "Observações"
      ]];

      sortedMaterialsForExcel().forEach(item => rows.push([
        item.code,
        item.description,
        item.size || "",
        item.manufacturer || "",
        item.category,
        item.ca,
        item.caValidity,
        item.lot,
        item.lotValidity,
        Number(item.stock),
        hasConfiguredMinimum(item) ? Number(item.minimum) : "",
        hasLifetime(item.lifetime) ? Number(item.lifetime) : "",
        item.location || "",
        item.notes || ""
      ]));

      const sheet = XLSX.utils.aoa_to_sheet(rows);
      sheet["!autofilter"] = { ref: sheet["!ref"] };
      sheet["!freeze"] = { xSplit: 0, ySplit: 1 };
      sheet["!cols"] = [
        { wch: 15 }, { wch: 36 }, { wch: 16 }, { wch: 26 }, { wch: 27 },
        { wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 19 }, { wch: 12 },
        { wch: 18 }, { wch: 20 }, { wch: 24 }, { wch: 40 }
      ];

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, "MATERIAIS");
      XLSX.writeFile(workbook, `materiais_epi_sese_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast("Planilha de materiais exportada em ordem alfabética e por tamanho.");
    } catch (error) {
      console.error("Falha na exportação Excel de materiais", error);
      showToast("Não foi possível gerar a planilha de materiais.", true);
    }
  }

  function interceptButton(button, label, handler) {
    if (!button) return;
    button.textContent = label;
    button.addEventListener("click", event => {
      event.preventDefault();
      event.stopImmediatePropagation();
      handler();
    }, true);
  }

  interceptButton(document.getElementById("exportMaterialsPdf"), "Relatório Excel", exportStockReportExcelPatched);
  interceptButton(document.getElementById("exportMaterials"), "Exportar Excel", exportMaterialsSortedPatched);
})();
