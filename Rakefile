# frozen_string_literal: true

require_relative 'lib/ooxml_parser'
require 'bundler/gem_tasks'
require 'rspec/core/rake_task'

RSpec::Core::RakeTask.new(:spec)

task default: :spec

desc 'Task for parse all files in directory'
task :parse_files, [:dir] do |_, args|
  files = Dir["#{args[:dir]}/**/*"].sort
  files.each do |file|
    next if File.directory? file

    p "Parsing file: #{file}"
    OoxmlParser::Parser.parse(file)
  end
end

desc 'Release gem '
task :release_github_rubygems do
  Rake::Task['release'].invoke
  gem_name = "pkg/#{OoxmlParser::Name::STRING}-"\
             "#{OoxmlParser::Version::STRING}.gem"
  sh('gem push --key github '\
     '--host https://rubygems.pkg.github.com/onlyoffice '\
     "#{gem_name}")
end
