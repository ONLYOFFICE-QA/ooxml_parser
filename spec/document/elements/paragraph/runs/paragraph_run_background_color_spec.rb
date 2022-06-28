# frozen_string_literal: true

require 'spec_helper'

describe 'ParagraphRun#background_color' do
  let(:docx) do
    OoxmlParser::Parser.parse('spec/document/elements/paragraph/' \
                              'runs/background_color/background_color.docx')
  end

  let(:background_color) { docx.elements.first.nonempty_runs.first.background_color }

  it 'Background color is a specific color' do
    expect(background_color).to eq(OoxmlParser::Color.new(226, 239, 216))
  end
end
