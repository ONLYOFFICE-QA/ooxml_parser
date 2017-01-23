require_relative 'lib/ooxml_parser'
require 'rspec/core/rake_task'

RSpec::Core::RakeTask.new(:spec)

task default: :spec

desc 'Task for parse all files in directory'
task :parse_files, [:dir] do |_, args|
  files = Dir["#{args[:dir]}/**/*"].shuffle
  files.each do |file|
    next if File.directory? file
    p "Parsing file: #{file}"
    OoxmlParser::Parser.parse(file)
  end
end
