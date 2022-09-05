# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  folder = 'spec/document/elements/paragraph/runs/object/ole_object/'

  it 'ole_object_file_reference_content' do
    docx = OoxmlParser::Parser.parse("#{folder}/spreadsheet_ole_object.docx")
    expect(docx.elements.first.nonempty_runs
               .first.object.ole_object
               .file_reference.content
               .worksheets[0].rows[0]
               .cells[0].text).to eq('1')
  end

  it 'spreadsheet_ole_and_link' do
    docx = OoxmlParser::Parser.parse("#{folder}/spreadsheet_ole_and_link.docx")
    expect(docx).to be_with_data
  end
end
