# frozen_string_literal: true

require 'spec_helper'

describe 'Group Char' do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_group_char.docx') }
  let(:group_char) { docx.elements.first.nonempty_runs.first.formula_run[1] }

  it 'symbol value' do
    expect(group_char.symbol).to eq('⏞')
  end

  it 'position value' do
    expect(group_char.position).to eq('top')
  end

  it 'vertical align value' do
    expect(group_char.vertical_align).to eq('bot')
  end
end
