# frozen_string_literal: true

require 'spec_helper'

describe 'Formula Accent' do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_accent.docx') }
  let(:accent) { docx.elements.first.nonempty_runs.first.formula_run[1] }

  it 'formula_accent parsed correctly' do
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'accent formula have correct symbol' do
    expect(accent.symbol).to eq("̇"v)
  end
end
