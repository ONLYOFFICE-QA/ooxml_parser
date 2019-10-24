# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'youtube_ole_object_id' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/object/ole_object/youtube_ole_object_id.docx')
    expect(docx.elements.first.nonempty_runs.first.object.ole_object.id).to eq('rId7')
  end

  it 'youtube_ole_object_file_reference_content' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/object/ole_object/youtube_ole_object_file_reference_content.docx')
    expect(docx.elements.first.nonempty_runs.first.object.ole_object.file_reference.content).to include('https://www.youtube.com/watch?v=D49vvl7BPro')
  end
end
