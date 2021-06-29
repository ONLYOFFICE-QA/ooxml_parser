# frozen_string_literal: true

require 'spec_helper'

describe 'DocxParagraphRun#hyperlink' do
  let(:hyperlink_file) do
    OoxmlParser::Parser.parse('spec/document/elements/'\
                              'paragraph/runs/'\
                              'pararga_run_hyperlink/'\
                              'paragraph_run_hyperlink.docx')
  end

  it 'hyperlink is not empty' do
    expect(hyperlink_file.elements.first
               .rows.first.cells.first
               .elements.first
               .character_style_array[1]
               .link.url).to include('http://www')
  end
end
