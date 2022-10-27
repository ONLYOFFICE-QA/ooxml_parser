# frozen_string_literal: true

require 'spec_helper'

describe 'Formula Bar' do
  let(:docx1) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_bar_1.docx') }
  let(:docx2) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_bar_2.docx') }

  it 'formula_bar_1.docx' do
    expect(docx1.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'formula_bar_2.docx' do
    expect(docx2.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'bar have position' do
    expect(docx2.elements.first.nonempty_runs.first.formula_run[1].position).to eq(:top)
  end
end
