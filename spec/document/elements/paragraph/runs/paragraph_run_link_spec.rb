# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Hyperlink do
  base_file_folder = 'spec/document/elements/paragraph/runs/link'
  it 'Check Hyperlink' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/hyperlink.docx")
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
  end

  it 'Check Hyperlink with tooltip' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/hyperlink_tooltip.docx")
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'LinkXpathError' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/link_xpath_error.docx")
    expect(docx.elements.first.nonempty_runs.first.text).to eq('yandex')
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'field_inside_hyperlink' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/field_inside_hyperlink.docx")
    expect(docx.elements.first.nonempty_runs.first.link.url).to eq('https://www.yandex.ru/')
  end

  it 'run_with_link' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/run_with_link.docx")
    expect(docx.elements.first.character_style_array[0].link.link).to eq('http://www.yandex.ru')
    expect(docx.elements.first.character_style_array[0].link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'hyperlink_with_run_into_it' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/hyperlink_with_run_into_it.docx")
    expect(docx).to be_with_data
  end

  it 'hyperlink runs text not merged' do
    docx = OoxmlParser::Parser.parse("#{base_file_folder}/hyperlink_runs_text_not_merged.docx")
    runs = docx.elements.first.sdt_content.paragraphs.first.nonempty_runs
    expect(runs[0].text).to eq('Simple Test Text')
    expect(runs[1].text).to eq("\t")
    expect(runs[2].text).to eq('1')
  end
end
