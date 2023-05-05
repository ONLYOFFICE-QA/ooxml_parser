# frozen_string_literal: true

require 'spec_helper'

describe 'Formula Accent' do
  docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_accent.docx')
  accent = docx.elements.first.nonempty_runs.first.formula_run[1]

  it 'formula_accent parsed correctly' do
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'accent formula have correct symbol' do
    expect(accent.symbol).to eq('̇')
  end
end
