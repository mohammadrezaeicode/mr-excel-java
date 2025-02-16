# MR Excel

MR Excel is a tool for generating Excel files. The implementation of this library is similar to ["`MR Excel JavaScript`"](https://github.com/mohammadrezaeicode/mr-excel-repo). You can use the documentation from that repository for reference. [example](https://github.com/mohammadrezaeicode/mr-excel-java-example-gallery) Requirement:**`JDK 11+`**

## Related Projects

The following list includes new repositories related to this project. Documentation and improvements for these projects can be found in the repositories below.

- **`MR Excel Java`**:A similar project using Java is in development. The release version is coming soon; currently, it is available as a snapshot version.["`repository`"](https://github.com/mohammadrezaeicode/mr-excel-java)

- **`MR Excel Editor`**: An editor that utilizes the library is currently under development. At present, it only generates simple results.["`repository`"](https:///github.com/mohammadrezaeicode/mr-excel-editor)["`Demo`"](https://mohammadrezaeicode.github.io/mr-excel-editor/)

## Import in `Maven` project

To use the dependency, follow the structure below:

1. `Generate a GitHub token` with access to `read:packages` and add the server to `~/.m2/setting.xml`. You can find the exact structure in this [link](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-apache-maven-registry#authenticating-with-a-personal-access-token).

```xml
<servers>
    <server>
        <id></id>
        <username></username>
        <password></password>
    </server>
</servers>
```

2. Add the repository to your `pom.xml` file:

```xml
<repositories>
  <repository>
    <id>github</id>
    <url>https://maven.pkg.github.com/mohammadrezaeicode/excel</url>
    <releases>
      <enabled>true</enabled>
    </releases>
    <snapshots>
      <enabled>true</enabled>
    </snapshots>
  </repository>
</repositories>
```

3. Add the library to your `dependencies`:

```xml
<dependency>
  <groupId>io.github.mohammadrezaeicode</groupId>
  <artifactId>excel</artifactId>
  <version>0.1-SNAPSHOT</version>
</dependency>
```

## Import in `Gradle` project

To use the dependency, follow the structure below:

1. **Generate a GitHub token** with access to `read:packages` and configure your `~/.gradle/gradle.properties` file:

```properties
gpr.user=YOUR_GITHUB_USERNAME
gpr.token=YOUR_GITHUB_TOKEN
```

2. **Add the repository to your `build.gradle` file**:

```groovy
repositories {
    maven {
        url = uri("https://maven.pkg.github.com/mohammadrezaeicode/excel")
        credentials {
            username = project.findProperty("gpr.user") ?: System.getenv("USERNAME")
            password = project.findProperty("gpr.token") ?: System.getenv("TOKEN")
        }
    }
}
```

3. **Add the library to your dependencies**:

```groovy
dependencies {
    implementation 'io.github.mohammadrezaeicode:excel:0.1-SNAPSHOT'
}
```


