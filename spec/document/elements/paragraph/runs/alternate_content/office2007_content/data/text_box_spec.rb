require 'spec_helper'

describe 'My behaviour' do
  it 'table_in_text_box' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/data/text_box/table_in_text_box.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content.office2007_content.data.text_box[1].rows.size).to eq(5)
    expect(docx.elements.first.nonempty_runs.first.alternate_content.office2007_content.data.text_box[1].rows.first.cells.size).to eq(5)
  end
end
