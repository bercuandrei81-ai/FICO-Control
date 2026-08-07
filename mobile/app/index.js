import React, { useEffect, useState } from "react";
import {
  SafeAreaView,
  View,
  Text,
  TextInput,
  Pressable,
  ActivityIndicator,
  Alert,
  StyleSheet,
  ScrollView
} from "react-native";
import { Picker } from "@react-native-picker/picker";
import { StatusBar } from "expo-status-bar";

// IMPORTANT:
// După ce backend-ul este online, înlocuiește adresa de mai jos
// cu URL-ul real, de ex. https://fico-control.onrender.com
const API_BASE = "http://127.0.0.1:8000";

export default function Home() {
  const [drivers, setDrivers] = useState([]);
  const [driverId, setDriverId] = useState("");
  const [score, setScore] = useState("");
  const [loading, setLoading] = useState(true);
  const [sending, setSending] = useState(false);

  const loadDrivers = async () => {
    try {
      setLoading(true);
      const res = await fetch(`${API_BASE}/api/drivers/today`);
      if (!res.ok) throw new Error("Nu s-a putut încărca lista.");
      const data = await res.json();
      setDrivers(data.drivers || []);
    } catch (e) {
      Alert.alert(
        "Conexiune",
        "Aplicația nu se poate conecta încă la server. După publicarea backend-ului online, aici va funcționa de oriunde."
      );
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadDrivers();
  }, []);

  const submit = async () => {
    if (!driverId) {
      Alert.alert("Lipsește numele", "Selectează numele șoferului.");
      return;
    }

    const fico = Number(score);
    if (!Number.isInteger(fico) || fico < 0 || fico > 1000) {
      Alert.alert("Scor invalid", "Introdu un scor FICO valid.");
      return;
    }

    try {
      setSending(true);
      const res = await fetch(`${API_BASE}/api/submissions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          driver_id: Number(driverId),
          fico_score: fico
        })
      });

      const data = await res.json().catch(() => ({}));

      if (!res.ok) {
        if (data.detail === "already_sent") {
          Alert.alert("Deja trimis", "Ai trimis deja scorul FICO pentru astăzi.");
        } else {
          Alert.alert("Eroare", "Scorul nu a putut fi trimis.");
        }
        return;
      }

      Alert.alert("Trimis", "Scorul FICO a fost înregistrat cu succes.");
      setScore("");
    } catch (e) {
      Alert.alert("Conexiune", "Nu s-a putut contacta serverul.");
    } finally {
      setSending(false);
    }
  };

  return (
    <SafeAreaView style={styles.safe}>
      <StatusBar style="dark" />
      <ScrollView contentContainerStyle={styles.page}>
        <View style={styles.card}>
          <Text style={styles.brand}>FICO CONTROL</Text>
          <Text style={styles.title}>Trimite scorul FICO</Text>
          <Text style={styles.subtitle}>
            Selectează numele și introdu scorul de astăzi.
          </Text>

          <Text style={styles.label}>Numele șoferului</Text>

          <View style={styles.pickerBox}>
            {loading ? (
              <ActivityIndicator style={{ padding: 16 }} />
            ) : (
              <Picker
                selectedValue={driverId}
                onValueChange={(v) => setDriverId(v)}
              >
                <Picker.Item label="Selectează numele" value="" />
                {drivers.map((d) => (
                  <Picker.Item key={d.id} label={d.name} value={String(d.id)} />
                ))}
              </Picker>
            )}
          </View>

          <Text style={styles.label}>Scor FICO</Text>
          <TextInput
            value={score}
            onChangeText={setScore}
            keyboardType="number-pad"
            placeholder="ex. 850"
            style={styles.input}
          />

          <Pressable
            onPress={submit}
            disabled={sending}
            style={({ pressed }) => [
              styles.button,
              pressed && { opacity: 0.85 },
              sending && { opacity: 0.6 }
            ]}
          >
            <Text style={styles.buttonText}>
              {sending ? "Se trimite..." : "Trimite scorul"}
            </Text>
          </Pressable>

          <Pressable onPress={loadDrivers}>
            <Text style={styles.refresh}>Actualizează lista</Text>
          </Pressable>
        </View>
      </ScrollView>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safe: { flex: 1, backgroundColor: "#f4f6f8" },
  page: { flexGrow: 1, justifyContent: "center", padding: 22 },
  card: {
    backgroundColor: "#ffffff",
    borderRadius: 22,
    padding: 24
  },
  brand: {
    fontSize: 13,
    fontWeight: "800",
    letterSpacing: 2.2,
    marginBottom: 18,
    color: "#17212b"
  },
  title: {
    fontSize: 30,
    fontWeight: "800",
    color: "#17212b",
    marginBottom: 8
  },
  subtitle: {
    fontSize: 15,
    color: "#667085",
    marginBottom: 25,
    lineHeight: 21
  },
  label: {
    fontSize: 16,
    fontWeight: "700",
    color: "#17212b",
    marginBottom: 8,
    marginTop: 14
  },
  pickerBox: {
    borderWidth: 1,
    borderColor: "#d8dde3",
    borderRadius: 12,
    overflow: "hidden",
    backgroundColor: "#fff"
  },
  input: {
    borderWidth: 1,
    borderColor: "#d8dde3",
    borderRadius: 12,
    paddingHorizontal: 15,
    paddingVertical: 14,
    fontSize: 18
  },
  button: {
    marginTop: 24,
    backgroundColor: "#17212b",
    padding: 16,
    borderRadius: 12,
    alignItems: "center"
  },
  buttonText: {
    color: "#fff",
    fontSize: 17,
    fontWeight: "800"
  },
  refresh: {
    textAlign: "center",
    marginTop: 18,
    fontWeight: "700",
    color: "#475467"
  }
});
