# frozen_string_literal: true

require 'spec_helper'

describe 'File name tests' do
  it 'Correct parse file by it full file name' do
    file_path = File.expand_path('spec/common/file_path/file_path.docx')
    docx = OoxmlParser::DocxParser.parse_docx(file_path)
    expect(docx).to be_a(OoxmlParser::DocumentStructure)
  end
end
