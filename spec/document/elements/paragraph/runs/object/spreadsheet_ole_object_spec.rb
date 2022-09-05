# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ole_object_file_reference_content' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/object/ole_object/spreadsheet_ole_object.docx')
    expect(docx.elements.first.nonempty_runs.first.object.ole_object.file_reference.content.worksheets[0].rows[0].cells[0].text).to eq('1')
  end
end
