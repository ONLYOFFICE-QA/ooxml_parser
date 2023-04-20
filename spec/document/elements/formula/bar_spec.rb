# frozen_string_literal: true

require 'spec_helper'

describe 'Formula Bar' do
  let(:first_docx) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_bar_1.docx') }
  let(:second_docx) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_bar_2.docx') }

  it 'formula_bar_1.docx' do
    expect(first_docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'formula_bar_2.docx' do
    expect(second_docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'bar have position' do
    expect(second_docx.elements.first.nonempty_runs.first.formula_run[1].position).to eq(:top)
  end
end
