# frozen_string_literal: true

require 'spec_helper'

describe 'DocxParagraphRun#instruction_text' do
  it 'DocxParagraphRun#instruction_text contains some text' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/paragraph_run_instruction_text/instruction_text_sample.docx')
    expect(docx.elements[1].runs[1].instruction_text.value).to eq(' REF _Ref1  \\h')
  end
end
