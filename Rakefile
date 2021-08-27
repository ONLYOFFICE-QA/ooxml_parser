# frozen_string_literal: true

require_relative 'lib/ooxml_parser'
require 'bundler'
Bundler::GemHelper.install_tasks(name: 'ooxml_parser')
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

desc 'Build windows gem'
task :build_mingw_gem do
  `gem build ooxml_parser-mingw32.gemspec`
  `mkdir -pv pkg`
  `mv *.gem pkg`
end

desc 'Release gem'
task release_github_rubygems: :build_mingw_gem do
  Rake::Task['release'].invoke
  default_gem = "pkg/#{OoxmlParser::Name::STRING}-"\
                "#{OoxmlParser::Version::STRING}.gem"
  sh('gem push --key github '\
     '--host https://rubygems.pkg.github.com/onlyoffice '\
     "#{default_gem}")
  mingw_gem = "pkg/#{OoxmlParser::Name::STRING}-"\
              "#{OoxmlParser::Version::STRING}-mingw32.gem"
  `gem push #{mingw_gem}`
end
