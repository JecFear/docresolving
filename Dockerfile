FROM java:8
VOLUME /tmp
COPY target/doc-resolving.jar app.jar
EXPOSE 9227
RUN bash -c 'touch /app.jar'
ENTRYPOINT ["java","-Djava.security.egd=file:/dev/./urandom","-jar","/app.jar"]
