require 'spec_helper'

describe OoxmlParser::Chart do
  it 'Header: insert chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/type/header_insert_chart.docx')
    elements = docx.page_properties.notes.first.elements
    expect(elements.first.character_style_array.first.drawing.graphic.data.type).to eq(:column)
    expect(elements.first.character_style_array.first.drawing.graphic.data.grouping).to eq(:clustered)
  end

  it 'Check Docx With Chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/type/docx_with_chart.docx')
    elements = docx.elements[1].rows.first.cells[0].elements
    expect(elements.first.nonempty_runs.first.drawing.graphic.data.type).to eq(:line)
    expect(elements.first.nonempty_runs.first.drawing.graphic.data.grouping).to eq(:standard)
  end

  it 'Check Docx With Several Chart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/type/docx_with_several_chart.docx')
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.grouping).to eq(:clustered)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.type).to eq(:column)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.grouping).to eq(:stacked)
  end

  it 'Different Pie Chats' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/type/pie_drawing_type.docx')
    expect(docx.elements[0].nonempty_runs.first.drawing.graphic.data.type).to eq(:pie)
    expect(docx.elements[1].nonempty_runs.first.drawing.graphic.data.type).to eq(:doughnut)
  end
end
